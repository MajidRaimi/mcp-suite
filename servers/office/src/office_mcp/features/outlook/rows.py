from __future__ import annotations

import re
from collections import OrderedDict
from collections.abc import Callable, Sequence
from dataclasses import dataclass, replace
from datetime import date, timedelta
from enum import StrEnum
from typing import TYPE_CHECKING

from office_mcp.features.outlook.listing import (
    has_load_more,
    parse_folder,
    parse_folders,
    parse_rows,
    pick_folder_row,
)
from office_mcp.features.outlook.models import Row
from office_mcp.features.outlook.window import (
    FocusedTab,
    OutlookMode,
    OutlookModule,
    folder_from_title,
)
from office_mcp.shared.errors import (
    AmbiguousTargetError,
    TargetError,
    TargetNotFoundError,
    UiTimeoutError,
)
from office_mcp.shared.focus import FrontHook
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

if TYPE_CHECKING:
    from office_mcp.features.outlook.gateway import OutlookGateway

KEY_PATTERN = re.compile(r"^[0-9a-f]{12}$")
DEFAULT_FOLDER = "Inbox"
DEFAULT_CAPACITY = 500
MAX_LOAD_MORE = 20
REFRESH_PAGES = 5
MAX_CANDIDATES = 6
SEARCH_SETTLE_ROUNDS = 2
TOP_RESULTS = normalize("Top Results")
KEYS_ONLY = "use the key from outlook_conversations or outlook_search"
KEY_GONE = "the conversation changed or a day passed; list or search again"


class SearchScope(StrEnum):
    CURRENT_FOLDER = "Current Folder"
    SUBFOLDERS = "Subfolders"
    MAILBOX = "Current Mailbox"
    ALL = "All Mailboxes"


@dataclass(frozen=True)
class SearchState:
    query: str
    scope: SearchScope


@dataclass(frozen=True)
class FolderLocator:
    folder: str
    view: FocusedTab | None = None


@dataclass(frozen=True)
class SearchLocator:
    query: str
    scope: SearchScope


Locator = FolderLocator | SearchLocator


def is_key(ref: str) -> bool:
    return KEY_PATTERN.match(ref) is not None


class RefRegistry:
    def __init__(self, capacity: int = DEFAULT_CAPACITY) -> None:
        self._capacity = max(capacity, 1)
        self._entries: OrderedDict[str, Locator] = OrderedDict()

    def remember(self, rows: Sequence[Row], locator: Locator) -> None:
        for row in rows:
            self._entries[row.summary.key] = locator
            self._entries.move_to_end(row.summary.key)
        while len(self._entries) > self._capacity:
            self._entries.popitem(last=False)

    def locate(self, key: str) -> Locator | None:
        found = self._entries.get(key)
        if found is not None:
            self._entries.move_to_end(key)
        return found


def open_folder(gateway: OutlookGateway, waiter: Waiter, name: str) -> None:
    wanted = normalize(name)
    if wanted and normalize(folder_from_title(gateway.window_title())) == wanted:
        return
    nodes = gateway.folder_rows()
    node = pick_folder_row(nodes, name)
    if node is None:
        known = ", ".join(dict.fromkeys(folder.name for folder in parse_folders(nodes)))
        raise TargetNotFoundError(f"no folder named {name!r}. Folders: {known}")
    target = parse_folder(node.description)
    shown = normalize(target.name) if target is not None else wanted
    gateway.select(node)
    waiter.require(
        lambda: normalize(folder_from_title(gateway.window_title())).startswith(shown),
        lambda: UiTimeoutError(f"Outlook did not open the folder {name!r} in time"),
    )


def outside_top_results(rows: Sequence[Row]) -> Row | None:
    groups = [normalize(row.summary.group) for row in rows]
    if TOP_RESULTS not in groups or len(set(groups)) != len(groups):
        return None
    rest = [row for row, group in zip(rows, groups, strict=True) if group != TOP_RESULTS]
    return rest[0] if len(rest) == 1 else None


def single_by_key(ref: str, rows: Sequence[Row]) -> Row:
    if len(rows) == 1:
        return rows[0]
    listed = outside_top_results(rows)
    if listed is None:
        raise candidates_error(ref, rows)
    return listed


def candidates_error(ref: str, rows: Sequence[Row]) -> AmbiguousTargetError:
    titles = [row.summary.title[:60] for row in rows[:MAX_CANDIDATES]]
    return AmbiguousTargetError(f"{ref!r} matches several conversations", titles)


class RowFinder:
    def __init__(
        self,
        gateway: OutlookGateway,
        waiter: Waiter,
        today: Callable[[], date],
        registry: RefRegistry,
        front: FrontHook | None = None,
    ) -> None:
        self._gateway = gateway
        self._waiter = waiter
        self._today = today
        self._registry = registry
        self._front = front or FrontHook()

    def rows(self, locator: Locator, days: int | None = None, load_more: int = 0) -> list[Row]:
        view = self._show(locator)
        nodes = self._load(days, load_more)
        return self._remember(parse_rows(nodes, self._today()), locator, view)

    def open_folder(self, name: str) -> None:
        open_folder(self._gateway, self._waiter, name)

    def current_locator(self) -> Locator:
        search = self._gateway.search_state()
        if search is not None:
            return SearchLocator(search.query, search.scope)
        folder = folder_from_title(self._gateway.window_title()) or DEFAULT_FOLDER
        return FolderLocator(folder, self._gateway.focused_tab())

    def locator_for(self, key: str, fallback_folder: str | None) -> Locator:
        known = self._registry.locate(key)
        if known is not None:
            return known
        if fallback_folder:
            return FolderLocator(fallback_folder)
        return self.current_locator()

    def resolve_for_read(self, ref: str, rows: list[Row]) -> Row:
        by_key = [row for row in rows if row.summary.key == ref]
        if by_key:
            return single_by_key(ref, by_key)
        if ref.isdigit():
            by_index = [row for row in rows if row.summary.index == int(ref)]
            if by_index:
                return by_index[0]
        wanted = normalize(ref)
        by_text = [row for row in rows if wanted and wanted in normalize(row.summary.title)]
        if len(by_text) == 1:
            return by_text[0]
        if by_text:
            raise candidates_error(ref, by_text)
        raise TargetNotFoundError(f"no conversation matches {ref!r}; list conversations again")

    def resolve_for_write(self, key: str, rows: list[Row]) -> Row:
        if not is_key(key):
            raise TargetError(KEYS_ONLY)
        matches = [row for row in rows if row.summary.key == key]
        if not matches:
            raise TargetNotFoundError(KEY_GONE)
        return single_by_key(key, matches)

    def refresh(self, key: str, locator: Locator) -> Row:
        if not is_key(key):
            raise TargetError(KEYS_ONLY)
        view = self._show(locator)
        nodes = self._gateway.message_rows()
        for _ in range(REFRESH_PAGES):
            if self._has_key(nodes, key):
                break
            more = has_load_more(nodes)
            if more is None:
                break
            nodes = self._load_more(more, len(nodes))
        rows = self._remember(parse_rows(nodes, self._today()), locator, view)
        return self.resolve_for_write(key, rows)

    def _show(self, locator: Locator) -> FocusedTab | None:
        if isinstance(locator, SearchLocator):
            self._show_search(locator)
            return None
        self._leave_search_and_calendar()
        if locator.folder:
            self.open_folder(locator.folder)
        tab = self._gateway.focused_tab()
        if locator.view is not None and tab is not None and tab != locator.view:
            self._front.bring_front()
            self._gateway.switch_tab(locator.view)
            return locator.view
        return tab

    def _show_search(self, locator: SearchLocator) -> None:
        wanted = SearchState(locator.query, locator.scope)
        if self._gateway.search_state() == wanted:
            return
        self._gateway.run_search(locator.query, locator.scope)
        self._waiter.settle(
            lambda: len(self._gateway.message_rows()), minimum=1, stable_rounds=SEARCH_SETTLE_ROUNDS
        )

    def _leave_search_and_calendar(self) -> None:
        mode = self._gateway.mode()
        if mode is OutlookMode.SEARCH:
            self._gateway.cancel_search()
        elif mode is OutlookMode.CALENDAR:
            self._gateway.show_module(OutlookModule.MAIL)

    def _load(self, days: int | None, load_more: int) -> list[Node]:
        nodes = self._gateway.message_rows()
        rounds = load_more if days is None else MAX_LOAD_MORE
        for _ in range(max(rounds, 0)):
            if days is not None and self._covers(nodes, days):
                break
            more = has_load_more(nodes)
            if more is None:
                break
            nodes = self._load_more(more, len(nodes))
        return nodes

    def _covers(self, nodes: list[Node], days: int) -> bool:
        since = (self._today() - timedelta(days=days)).isoformat()
        dated = [row.summary.date for row in parse_rows(nodes, self._today()) if row.summary.date]
        return bool(dated) and min(dated) < since

    def _load_more(self, more: Node, count: int) -> list[Node]:
        self._gateway.press(more)
        grown = self._waiter.until(lambda: self._grown(count))
        return grown if grown is not None else self._gateway.message_rows()

    def _grown(self, count: int) -> list[Node] | None:
        nodes = self._gateway.message_rows()
        return nodes if len(nodes) > count else None

    def _has_key(self, nodes: list[Node], key: str) -> bool:
        return any(row.summary.key == key for row in parse_rows(nodes, self._today()))

    def _remember(self, rows: list[Row], locator: Locator, view: FocusedTab | None) -> list[Row]:
        labelled = [label_row(row, locator, view) for row in rows]
        self._registry.remember(labelled, locator)
        return labelled


def label_row(row: Row, locator: Locator, view: FocusedTab | None) -> Row:
    if isinstance(locator, SearchLocator):
        return row
    summary = replace(
        row.summary,
        folder=row.summary.folder or locator.folder,
        view=view.value if view is not None else row.summary.view,
    )
    return Row(summary=summary, node=row.node)
