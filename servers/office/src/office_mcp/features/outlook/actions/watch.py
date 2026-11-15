from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date

from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.listing import parse_rows
from office_mcp.features.outlook.models import Row, Summary
from office_mcp.features.outlook.rows import Locator
from office_mcp.shared.node import Node
from office_mcp.shared.waiting import Waiter

COMPLETED = "completed"
GONE_WAIT = 5.0
STATE_WAIT = 4.0
SELECT_WAIT = 3.0
MULTI_SELECTION = re.compile(
    r"\b\d+\s+(?:items|conversations|messages)\s+selected\b|\bselected\s+\d+\s+items\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Target:
    key: str
    locator: Locator
    row: Row

    @property
    def label(self) -> str:
        return self.row.summary.title[:80]

    @property
    def folder(self) -> str:
        return self.row.summary.folder


class RowWatch:
    def __init__(self, outlook: OutlookGateway, waiter: Waiter, today: Callable[[], date]) -> None:
        self._outlook = outlook
        self._waiter = waiter
        self._today = today

    def current(self) -> list[Row]:
        return parse_rows(self._outlook.message_rows(), self._today())

    def find(self, key: str) -> Row | None:
        return next((row for row in self.current() if row.summary.key == key), None)

    def selected_rows(self) -> list[Row]:
        return [row for row in self.current() if row.node.selected]

    def selected_keys(self) -> list[str]:
        return [row.summary.key for row in self.selected_rows()]

    def multi_selection_shown(self) -> bool:
        return shows_multi_selection(self._outlook.reading_pane())

    def wait_selected(self, key: str, timeout: float = SELECT_WAIT) -> bool:
        return bool(self._waiter.until(lambda: self.selected_keys() == [key], timeout))

    def wait_gone(self, key: str, timeout: float = GONE_WAIT) -> bool:
        return bool(self._waiter.until(lambda: self.find(key) is None, timeout))

    def wait_for(
        self, key: str, check: Callable[[Summary], bool], timeout: float = STATE_WAIT
    ) -> bool:
        return bool(self._waiter.until(lambda: self._matches(key, check), timeout))

    def _matches(self, key: str, check: Callable[[Summary], bool]) -> bool:
        row = self.find(key)
        return row is not None and check(row.summary)


def is_unread(summary: Summary) -> bool:
    return summary.unread


def is_read(summary: Summary) -> bool:
    return not summary.unread


def is_flagged(summary: Summary) -> bool:
    return summary.flagged


def is_unflagged(summary: Summary) -> bool:
    return not summary.flagged and COMPLETED not in summary.flags


def shows_multi_selection(pane: Node | None) -> bool:
    if pane is None:
        return False
    return any(
        MULTI_SELECTION.search(part)
        for node in pane.walk()
        for part in (node.title, node.value, node.description)
        if part
    )
