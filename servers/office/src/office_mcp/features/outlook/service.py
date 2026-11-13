from __future__ import annotations

import time
from collections.abc import Callable
from dataclasses import replace
from datetime import date, timedelta

from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.listing import has_load_more, parse_folder, parse_rows
from office_mcp.features.outlook.models import Email, Folder, Row, Summary
from office_mcp.features.outlook.reading import DEFAULT_MAX_BODY, parse_email
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize

POLL_SECONDS = 0.4
MARK_UNREAD = "Mark Unread"
MARK_READ = "Mark Read"
TITLE_SEPARATOR = " • "
MAX_LOAD_MORE = 20
BODY_GRACE_SECONDS = 4.0


class OutlookService:
    def __init__(
        self,
        gateway: OutlookGateway,
        today: Callable[[], date],
        timeout: float = 12.0,
        sleep: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self._gateway = gateway
        self._today = today
        self._timeout = timeout
        self._sleep = sleep
        self._clock = clock

    def folders(self) -> list[Folder]:
        parsed = (parse_folder(node.description) for node in self._gateway.folder_rows())
        return [folder for folder in parsed if folder]

    def conversations(
        self,
        folder: str | None,
        unread_only: bool,
        days: int | None,
        limit: int,
        load_more: int = 0,
    ) -> list[Summary]:
        rows = self._rows(folder, days, load_more)
        summaries = [row.summary for row in rows]
        if unread_only:
            summaries = [summary for summary in summaries if summary.unread]
        if days is not None:
            since = (self._today() - timedelta(days=days)).isoformat()
            summaries = [s for s in summaries if s.date is not None and s.date >= since]
        return summaries[:limit] if limit > 0 else summaries

    def read(
        self,
        ref: str,
        folder: str | None,
        keep_unread: bool = True,
        max_body: int = DEFAULT_MAX_BODY,
    ) -> Email:
        row = self._resolve(ref, self._rows(folder, None, 0))
        self._gateway.select(row.node)
        email = self._wait_for_email(row.summary, max_body)
        if keep_unread and row.summary.unread:
            self._mark_unread()
        return replace(email, summary=row.summary)

    def _rows(self, folder: str | None, days: int | None, load_more: int) -> list[Row]:
        if folder:
            self._open_folder(folder)
        nodes = self._gateway.message_rows()
        rounds = load_more if days is None else MAX_LOAD_MORE
        for _ in range(max(rounds, 0)):
            if days is not None and self._covers(nodes, days):
                break
            more = has_load_more(nodes)
            if more is None:
                break
            nodes = self._load_more(more, len(nodes))
        return parse_rows(nodes, self._today())

    def _covers(self, nodes: list[Node], days: int) -> bool:
        since = (self._today() - timedelta(days=days)).isoformat()
        dated = [row.summary.date for row in parse_rows(nodes, self._today()) if row.summary.date]
        return bool(dated) and min(dated) < since

    def _load_more(self, more: Node, count: int) -> list[Node]:
        self._gateway.press(more)
        grown = self._wait(lambda: self._grown(count))
        return grown if grown is not None else self._gateway.message_rows()

    def _grown(self, count: int) -> list[Node] | None:
        nodes = self._gateway.message_rows()
        return nodes if len(nodes) > count else None

    def _open_folder(self, name: str) -> None:
        wanted = normalize(name)
        if normalize(self._gateway.window_title().split(TITLE_SEPARATOR, 1)[0]) == wanted:
            return
        rows = self._gateway.folder_rows()
        folders = [(node, parse_folder(node.description)) for node in rows]
        matches = [node for node, folder in folders if folder and normalize(folder.name) == wanted]
        matches = matches or [
            node for node, folder in folders if folder and normalize(folder.name).startswith(wanted)
        ]
        if not matches:
            known = ", ".join(folder.name for _, folder in folders if folder)
            raise OfficeError(f"no folder named {name!r}. Folders: {known}")
        self._gateway.select(matches[0])
        opened = self._wait(lambda: self._title_is(wanted))
        if opened is None:
            raise OfficeError(f"Outlook did not open the folder {name!r} in time")

    def _title_is(self, wanted: str) -> bool | None:
        title = self._gateway.window_title().split(TITLE_SEPARATOR, 1)[0]
        return True if normalize(title).startswith(wanted) else None

    def _resolve(self, ref: str, rows: list[Row]) -> Row:
        by_key = [row for row in rows if row.summary.key == ref]
        if by_key:
            return by_key[0]
        if ref.isdigit():
            by_index = [row for row in rows if row.summary.index == int(ref)]
            if by_index:
                return by_index[0]
        wanted = normalize(ref)
        by_text = [row for row in rows if wanted in normalize(row.summary.title)]
        if len(by_text) == 1:
            return by_text[0]
        if by_text:
            titles = "; ".join(row.summary.title[:60] for row in by_text[:6])
            raise OfficeError(f"{ref!r} matches several conversations: {titles}")
        raise OfficeError(f"no conversation matches {ref!r}; list conversations again")

    def _wait_for_email(self, summary: Summary, max_body: int) -> Email:
        def ready() -> Email | None:
            pane = self._gateway.reading_pane()
            email = parse_email(pane, max_body) if pane else None
            if not email or not email.subject or not subject_fits(email, summary):
                return None
            return email if email.body or self._clock() >= body_deadline else None

        body_deadline = self._clock() + BODY_GRACE_SECONDS
        email = self._wait(ready)
        if email is None:
            raise OfficeError(f"Outlook did not show the conversation {summary.title[:60]!r}")
        return email

    def _mark_unread(self) -> None:
        button = self._wait(lambda: self._gateway.toolbar_button(MARK_UNREAD))
        if button is None:
            return
        self._gateway.press_button(button)
        if self._wait(lambda: self._gateway.toolbar_button(MARK_READ)) is None:
            raise OfficeError("read the email, but could not mark it unread again")

    def _wait[T](self, ready: Callable[[], T | None]) -> T | None:
        deadline = self._clock() + self._timeout
        while True:
            result = ready()
            if result is not None:
                return result
            if self._clock() >= deadline:
                return None
            self._sleep(POLL_SECONDS)


def subject_fits(email: Email, summary: Summary) -> bool:
    subject = normalize(email.subject)
    for prefix in ("re:", "fw:", "fwd:", "canceled:", "cancelled:", "accepted:", "declined:"):
        subject = subject.removeprefix(prefix).strip()
    return not subject or subject[:25] in normalize(summary.title)
