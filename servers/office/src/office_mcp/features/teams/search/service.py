from __future__ import annotations

from collections.abc import Callable
from contextlib import suppress
from datetime import date

from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.reading.view import ReadScope
from office_mcp.features.teams.search.models import SearchResult
from office_mcp.features.teams.search.parse import parse_results, search_input
from office_mcp.shared.errors import ElementNotFoundError, OfficeError, TargetError
from office_mcp.shared.events import Key, Shortcut
from office_mcp.shared.node import Node
from office_mcp.shared.waiting import Waiter

MIN_QUERY = 2
MAX_QUERY = 200
RESULT_WAIT = 6.0
SETTLE_WAIT = 4.0
ESCAPE = Shortcut(Key.ESCAPE)
CLEAR_ESCAPES = 2
CLEAR_SEARCH = "clear the Teams search box"


def checked_query(query: str) -> str:
    cleaned = " ".join(query.split())
    if not MIN_QUERY <= len(cleaned) <= MAX_QUERY:
        raise TargetError(f"the search text must be {MIN_QUERY} to {MAX_QUERY} characters long")
    return cleaned


class SearchReader:
    def __init__(self, gateway: TeamsGateway, waiter: Waiter, today: Callable[[], date]) -> None:
        self._gateway = gateway
        self._waiter = waiter
        self._today = today

    def search(self, scope: ReadScope, query: str, limit: int) -> list[SearchResult]:
        text = checked_query(query)
        box = search_input(self._gateway.snapshot())
        if box is None:
            raise ElementNotFoundError("the Teams search box was not found")
        scope.cleanup.push(CLEAR_SEARCH, lambda: self._clear(box))
        self._focus(box)
        self._gateway.set_value(box, text)
        if self._waiter.until(lambda: self._typed_results(text), RESULT_WAIT) is None:
            return []
        results = self._settled()
        return results[:limit] if limit > 0 else results

    def _focus(self, box: Node) -> None:
        with suppress(OfficeError):
            self._gateway.focus(box)

    def _typed_results(self, text: str) -> list[SearchResult] | None:
        root = self._gateway.snapshot()
        box = search_input(root)
        if box is not None and box.value and box.value != text:
            return None
        return parse_results(root, self._today()) or None

    def _settled(self) -> list[SearchResult]:
        latest: list[list[SearchResult]] = [[]]

        def sample() -> int:
            latest[0] = parse_results(self._gateway.snapshot(), self._today())
            return len(latest[0])

        self._waiter.settle(sample, timeout=SETTLE_WAIT)
        return latest[0]

    def _clear(self, box: Node) -> None:
        self._gateway.set_value(box, "")
        for _ in range(CLEAR_ESCAPES):
            self._gateway.press_key(ESCAPE)
