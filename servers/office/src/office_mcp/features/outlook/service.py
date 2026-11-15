from __future__ import annotations

import re
from collections.abc import Callable, Iterator, Sequence
from contextlib import contextmanager
from dataclasses import dataclass, replace
from datetime import date, datetime, timedelta
from enum import StrEnum
from typing import Any, Protocol

from office_mcp.features.outlook.conversation import LatestMessage
from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.health import OutlookHealth, read_health, unknown_health_record
from office_mcp.features.outlook.listing import (
    expanded_header_of,
    has_load_more,
    parse_folders,
    parse_rows,
)
from office_mcp.features.outlook.models import Email, Folder, Row, Summary
from office_mcp.features.outlook.reading import (
    DEFAULT_MAX_BODY,
    DEFAULT_MAX_MESSAGES,
    parse_email,
)
from office_mcp.features.outlook.rows import (
    MAX_LOAD_MORE,
    FolderLocator,
    Locator,
    RowFinder,
    is_key,
)
from office_mcp.features.outlook.thread import top_text
from office_mcp.features.outlook.window import FocusedTab
from office_mcp.shared.errors import (
    ActionNotVerifiedError,
    ElementNotFoundError,
    OfficeError,
    UiTimeoutError,
    error_text,
)
from office_mcp.shared.menus import MenuBar
from office_mcp.shared.node import Node
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import ScopeFactory, ScopeOptions, UiScope
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

MARK_AS_UNREAD = ("Message", "Mark as Unread")
TOOLBAR_MARK_UNREAD = "Mark Unread"
BODY_GRACE_SECONDS = 4.0
UNREAD_WAIT_SECONDS = 3.0
UNREAD_ATTEMPTS = 2
READ_SETTLE_SECONDS = 1.5
COLLAPSE_SECONDS = 2.0
SELECT_SECONDS = 2.0
FAVORITES = "favorites"
SUBJECT_PREFIXES = ("re:", "fw:", "fwd:", "canceled:", "cancelled:", "accepted:", "declined:")
CLOCK = re.compile(r"^(?P<hour>\d{1,2}):(?P<minute>\d{2})(?:\s*(?P<half>[AP]M))?$", re.IGNORECASE)
HALF_DAY_MINUTES = 12 * 60
LEAVE_CONVERSATION = "leave the conversation"


class MailView(StrEnum):
    CURRENT = "current"
    FOCUSED = "focused"
    OTHER = "other"
    BOTH = "both"


class UnreadMarker(Protocol):
    def mark_unread(self) -> None: ...


class MenuUnreadMarker:
    def __init__(self, menus: MenuBar) -> None:
        self._menus = menus

    def mark_unread(self) -> None:
        self._menus.press(MARK_AS_UNREAD)


class ToolbarUnreadMarker:
    def __init__(self, gateway: OutlookGateway) -> None:
        self._gateway = gateway

    def mark_unread(self) -> None:
        button = self._gateway.toolbar_button(TOOLBAR_MARK_UNREAD)
        if button is None:
            raise ElementNotFoundError("the Outlook Mark Unread button was not found")
        self._gateway.press_button(button)


@dataclass(frozen=True)
class OpenedEmail:
    row: Row
    locator: Locator
    email: Email
    pane: Node


class OpenConversation:
    def __init__(
        self,
        service: OutlookService,
        ui: UiScope,
        row: Row,
        locator: Locator,
        restore_unread: bool,
    ) -> None:
        self._service = service
        self._ui = ui
        self.row = row
        self.locator = locator
        self._restore_unread = restore_unread
        self._loaded: tuple[Email, Node] | None = None
        self._closed = False

    def load(self, email: Email, pane: Node) -> None:
        self._loaded = (email, pane)

    @property
    def opened(self) -> OpenedEmail:
        if self._loaded is None:
            raise UiTimeoutError(
                f"Outlook did not show the conversation {self.row.summary.title!r}"
            )
        email, pane = self._loaded
        return OpenedEmail(self.row, self.locator, email, pane)

    def close(self) -> tuple[str, ...]:
        if self._closed:
            return ()
        self._closed = True
        return self._service.leave(self._ui, self.row, self.locator, self._restore_unread)

    def close_or_raise(self) -> None:
        notes = self.close()
        if notes:
            raise ActionNotVerifiedError("; ".join(notes))


class OutlookService:
    def __init__(
        self,
        gateway: OutlookGateway,
        rows: RowFinder,
        scopes: ScopeFactory,
        waiter: Waiter,
        today: Callable[[], date],
        now: Callable[[], datetime],
        unread: UnreadMarker | None = None,
    ) -> None:
        self._gateway = gateway
        self._rows = rows
        self._scopes = scopes
        self._waiter = waiter
        self._today = today
        self._now = now
        self._unread = unread if unread is not None else ToolbarUnreadMarker(gateway)

    def health(self, launch: bool = True) -> OutlookHealth:
        return read_health(self._gateway.shallow_button_labels(launch), self._now())

    def health_record(self) -> dict[str, Any]:
        try:
            return self.health(launch=False).to_record()
        except OfficeError as error:
            return unknown_health_record(str(error))

    @contextmanager
    def scope(self, foreground: bool = False) -> Iterator[UiScope]:
        with self._scopes(ScopeOptions(foreground=foreground)) as ui:
            yield ui

    def folders(self) -> list[Folder]:
        with self.scope():
            return listed_folders(self._gateway.folder_rows())

    def conversations(
        self,
        folder: str | None,
        unread_only: bool,
        days: int | None,
        limit: int,
        load_more: int = 0,
        view: MailView = MailView.CURRENT,
        deadline: Deadline | None = None,
    ) -> list[Summary]:
        with self.scope():
            summaries = self.list_view(folder, view, days, load_more, deadline)
        if unread_only:
            summaries = [summary for summary in summaries if summary.unread]
        if days is not None:
            since = (self._today() - timedelta(days=days)).isoformat()
            summaries = [s for s in summaries if s.date is not None and s.date >= since]
        return summaries[:limit] if limit > 0 else summaries

    def list_view(
        self,
        folder: str | None,
        view: MailView,
        days: int | None,
        load_more: int,
        deadline: Deadline | None,
    ) -> list[Summary]:
        base = self._folder_locator(folder)
        if view is MailView.FOCUSED or view is MailView.OTHER:
            return self.list_rows(with_tab(base, FocusedTab(view.value)), days, load_more, deadline)
        first = self.list_rows(base, days, load_more, deadline)
        shown = self._gateway.focused_tab()
        if shown is None:
            return first
        pinned = self._pin_tab(base, shown, first)
        if view is MailView.CURRENT:
            return pinned
        second = self.list_rows(with_tab(base, other_tab(shown)), days, load_more, deadline)
        return merge_newest_first(pinned, second)

    def _pin_tab(self, base: Locator, shown: FocusedTab, listed: list[Summary]) -> list[Summary]:
        if not isinstance(base, FolderLocator) or base.view is not None:
            return listed
        return [row.summary for row in self._rows.rows(with_tab(base, shown))]

    def list_rows(
        self, locator: Locator, days: int | None, load_more: int, deadline: Deadline | None
    ) -> list[Summary]:
        first = self._rows.rows(locator)
        pages = MAX_LOAD_MORE if days is not None else load_more
        if pages <= 0 or (days is not None and covers(first, days, self._today())):
            return [row.summary for row in first]
        today = self._today()
        self.load_pages(
            pages, deadline, lambda nodes: days is not None and covers_nodes(nodes, days, today)
        )
        return [row.summary for row in self._rows.rows(locator)]

    def load_pages(
        self,
        pages: int,
        deadline: Deadline | None,
        enough: Callable[[list[Node]], bool] = lambda _nodes: False,
    ) -> int:
        nodes = self._gateway.message_rows()
        loaded = 0
        while loaded < pages and not expired(deadline) and not enough(nodes):
            more = has_load_more(nodes)
            if more is None:
                break
            nodes = self._load_page(more, len(nodes), deadline)
            loaded += 1
        return loaded

    def _load_page(self, more: Node, count: int, deadline: Deadline | None) -> list[Node]:
        self._gateway.press(more)
        grown = self._waiter.until(lambda: self._grown(count), remaining(deadline))
        return grown if grown is not None else self._gateway.message_rows()

    def read(
        self,
        ref: str,
        folder: str | None,
        keep_unread: bool = True,
        max_body: int = DEFAULT_MAX_BODY,
        max_messages: int = DEFAULT_MAX_MESSAGES,
    ) -> Email:
        with self.scope() as ui:
            conversation = self.open(ui, ref, folder, keep_unread, max_body, max_messages)
            notes = conversation.close()
        opened = conversation.opened
        return replace(opened.email, summary=opened.row.summary, notes=opened.email.notes + notes)

    def latest_message(self, key: str, folder: str | None) -> LatestMessage:
        with self.scope() as ui:
            conversation = self.open(ui, key, folder, True, DEFAULT_MAX_BODY, 1, keys_only=True)
            conversation.close_or_raise()
        return latest_from(conversation.opened)

    def open(
        self,
        ui: UiScope,
        ref: str,
        folder: str | None,
        keep_unread: bool,
        max_body: int,
        max_messages: int,
        keys_only: bool = False,
    ) -> OpenConversation:
        locator, row = self._locate(ref, folder, keys_only)
        self._gateway.select(row.node)
        restore = keep_unread and row.summary.unread
        conversation = OpenConversation(self, ui, row, locator, restore)
        ui.cleanup.push(LEAVE_CONVERSATION, conversation.close_or_raise)
        conversation.load(*self._wait_for_email(row.summary, max_body, max_messages))
        return conversation

    def leave(
        self, ui: UiScope, row: Row, locator: Locator, restore_unread: bool
    ) -> tuple[str, ...]:
        notes = self._collapse(ui, row.summary)
        if not restore_unread:
            self._gateway.deselect(row.node)
            return notes
        return (*notes, *self._restore_unread(ui, row.summary, locator))

    def _collapse(self, ui: UiScope, summary: Summary) -> tuple[str, ...]:
        if not self._shows_expanded(summary):
            return ()
        title = summary.title[:60]
        try:
            ui.ensure_front()
            self._gateway.collapse_selected()
            if self._waiter.until(lambda: not self._shows_expanded(summary), COLLAPSE_SECONDS):
                return ()
        except OfficeError as error:
            return (f"could not collapse {title!r} again: {error_text(error)}",)
        return (f"{title!r} was left expanded in the message list",)

    def _shows_expanded(self, summary: Summary) -> bool:
        rows = parse_rows(self._gateway.message_rows(), self._today())
        return any(expanded_header_of(row.summary, summary) for row in rows)

    def _restore_unread(self, ui: UiScope, summary: Summary, locator: Locator) -> tuple[str, ...]:
        title = summary.title[:60]
        try:
            ui.ensure_front()
            restored = self._keep_unread(summary.key, locator)
        except OfficeError as error:
            return (f"could not mark {title!r} unread again: {error_text(error)}",)
        if not restored:
            return (f"could not mark {title!r} unread again; it is now read",)
        return ()

    def _keep_unread(self, key: str, locator: Locator) -> bool:
        for attempt in range(UNREAD_ATTEMPTS):
            if attempt:
                self._reselect(key, locator)
            marked = self._mark_unread_verified(key, locator)
            self._gateway.deselect(self._rows.refresh(key, locator).node)
            if marked and not self._turns_read(key, locator):
                return True
        return False

    def _reselect(self, key: str, locator: Locator) -> None:
        self._gateway.select(self._rows.refresh(key, locator).node)
        self._waiter.until(lambda: self._rows.refresh(key, locator).node.selected, SELECT_SECONDS)

    def _turns_read(self, key: str, locator: Locator) -> bool:
        return bool(
            self._waiter.until(lambda: not self._is_unread(key, locator), READ_SETTLE_SECONDS)
        )

    def _mark_unread_verified(self, key: str, locator: Locator) -> bool:
        if self._is_unread(key, locator):
            return True
        self._unread.mark_unread()
        return bool(self._waiter.until(lambda: self._is_unread(key, locator), UNREAD_WAIT_SECONDS))

    def _is_unread(self, key: str, locator: Locator) -> bool:
        return self._rows.refresh(key, locator).summary.unread

    def _locate(self, ref: str, folder: str | None, keys_only: bool) -> tuple[Locator, Row]:
        if keys_only or is_key(ref):
            locator = self._rows.locator_for(ref, folder)
            return locator, self._rows.refresh(ref, locator)
        locator = self._folder_locator(folder)
        return locator, self._rows.resolve_for_read(ref, self._rows.rows(locator))

    def _folder_locator(self, folder: str | None) -> Locator:
        return FolderLocator(folder) if folder else self._rows.current_locator()

    def _grown(self, count: int) -> list[Node] | None:
        nodes = self._gateway.message_rows()
        return nodes if len(nodes) > count else None

    def _wait_for_email(
        self, summary: Summary, max_body: int, max_messages: int
    ) -> tuple[Email, Node]:
        clock = self._waiter.clock
        body_deadline = clock() + BODY_GRACE_SECONDS

        def ready() -> tuple[Email, Node] | None:
            pane = self._gateway.reading_pane()
            email = parse_email(pane, max_body, max_messages) if pane else None
            if pane is None or not email or not email.subject or not subject_fits(email, summary):
                return None
            return (email, pane) if email.body or clock() >= body_deadline else None

        return self._waiter.require(
            ready,
            lambda: UiTimeoutError(f"Outlook did not show the conversation {summary.title[:60]!r}"),
        )


def listed_folders(nodes: list[Node]) -> list[Folder]:
    return [
        folder
        for folder in parse_folders(nodes)
        if not folder.favorite and normalize(folder.path) != FAVORITES
    ]


def with_tab(locator: Locator, tab: FocusedTab) -> Locator:
    return replace(locator, view=tab) if isinstance(locator, FolderLocator) else locator


def other_tab(tab: FocusedTab) -> FocusedTab:
    return FocusedTab.OTHER if tab is FocusedTab.FOCUSED else FocusedTab.FOCUSED


def merge_newest_first(*groups: Sequence[Summary]) -> list[Summary]:
    seen: set[str] = set()
    merged: list[Summary] = []
    for summary in (item for group in groups for item in group):
        if summary.key not in seen:
            seen.add(summary.key)
            merged.append(summary)
    return sorted(merged, key=recency, reverse=True)


def recency(summary: Summary) -> tuple[str, int]:
    return (summary.date or "", clock_minutes(summary.time_label))


def clock_minutes(label: str) -> int:
    match = CLOCK.match(label.strip())
    if match is None:
        return -1
    hour, minute = int(match["hour"]), int(match["minute"])
    half = (match["half"] or "").upper()
    if half:
        hour = hour % 12 + (12 if half == "PM" else 0)
    minutes = hour * 60 + minute
    return minutes if minutes < 2 * HALF_DAY_MINUTES else -1


def covers(rows: Sequence[Row], days: int, today: date) -> bool:
    since = (today - timedelta(days=days)).isoformat()
    dated = [row.summary.date for row in rows if row.summary.date]
    return bool(dated) and min(dated) < since


def covers_nodes(nodes: list[Node], days: int, today: date) -> bool:
    return covers(parse_rows(nodes, today), days, today)


def expired(deadline: Deadline | None) -> bool:
    return deadline is not None and deadline.expired()


def remaining(deadline: Deadline | None) -> float | None:
    return None if deadline is None else deadline.remaining()


def latest_from(opened: OpenedEmail) -> LatestMessage:
    email = opened.email
    return LatestMessage(
        key=opened.row.summary.key,
        subject=email.subject,
        sender=email.sender,
        to=tuple(email.recipients.get("to", [])),
        cc=tuple(email.recipients.get("cc", [])),
        hidden_recipients=sum(email.hidden_recipients.values()),
        body_top=top_text(email.body),
        sent=email.sent,
        folder=opened.row.summary.folder,
    )


def subject_fits(email: Email, summary: Summary) -> bool:
    subject = normalize(email.subject)
    for prefix in SUBJECT_PREFIXES:
        subject = subject.removeprefix(prefix).strip()
    return not subject or subject[:25] in normalize(summary.title)
