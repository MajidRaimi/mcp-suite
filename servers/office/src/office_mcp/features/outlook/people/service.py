from __future__ import annotations

from collections import OrderedDict
from collections.abc import Callable, Iterable
from typing import Protocol

from office_mcp.features.outlook.mail.search import FIELD_DEPTH, search_field
from office_mcp.features.outlook.models import Person
from office_mcp.features.outlook.people.parse import (
    Contact,
    contact_from,
    matches_query,
    parse_suggestions,
    typed_rows,
    unique_contacts,
)
from office_mcp.features.outlook.suggestions import await_suggestions, close_suggestions
from office_mcp.shared.ax import AxApp
from office_mcp.shared.errors import ElementNotFoundError
from office_mcp.shared.events import InputEvents
from office_mcp.shared.node import Node
from office_mcp.shared.scope import ScopeFactory, ScopeOptions
from office_mcp.shared.tree import find
from office_mcp.shared.waiting import Waiter

SUGGESTION_WAIT_SECONDS = 3.0
SUGGESTION_MARK = "suggestion"
TABLE_DEPTH = 6
RECENT_CAPACITY = 300
CLEAR_FIELD = "clear the search field and close its suggestions"
PEOPLE_SCOPE = ScopeOptions(foreground=False)


class SuggestionSource(Protocol):
    def type(self, query: str) -> None: ...

    def rows(self) -> list[str]: ...

    def clear(self) -> None: ...


class AxSuggestions:
    def __init__(
        self, app: AxApp, window: Callable[[], object], events: InputEvents, waiter: Waiter
    ) -> None:
        self._app = app
        self._window = window
        self._events = events
        self._waiter = waiter

    def type(self, query: str) -> None:
        field = self._field()
        if field is None:
            raise ElementNotFoundError("the Outlook search field was not found")
        self._app.set_attribute(field, "AXFocused", True)
        self._app.set_attribute(field, "AXValue", query)

    def rows(self) -> list[str]:
        table = self._table()
        if table is None:
            return []
        snapshot = self._app.snapshot(table, max_depth=TABLE_DEPTH)
        return [row_text(node) for node in snapshot.walk() if node.role == "AXRow"]

    def clear(self) -> None:
        field = self._field()
        if field is None:
            return
        self._app.set_attribute(field, "AXValue", "")
        await_suggestions(self._app, self._waiter)
        close_suggestions(self._app, self._events, self._waiter)

    def _field(self) -> object | None:
        return search_field(self._app, self._window())

    def _table(self) -> object | None:
        field = self._field()
        if field is None:
            return None
        return find(
            self._app,
            field,
            lambda element: self._app.text(element, "AXRole") == "AXTable",
            FIELD_DEPTH,
        )


def row_text(row: Node) -> str:
    labels = [node.label() for node in row.walk()]
    return next((label for label in labels if SUGGESTION_MARK in label.casefold()), row.label())


class PeopleService:
    def __init__(self, source: SuggestionSource, scopes: ScopeFactory, waiter: Waiter) -> None:
        self._source = source
        self._scopes = scopes
        self._waiter = waiter
        self._recent: OrderedDict[str, Contact] = OrderedDict()

    def remember(self, people: Iterable[Person]) -> None:
        for person in people:
            contact = contact_from(person)
            if contact is None:
                continue
            address = contact.email.casefold()
            self._recent[address] = contact
            self._recent.move_to_end(address)
        while len(self._recent) > RECENT_CAPACITY:
            self._recent.popitem(last=False)

    def lookup(self, query: str, limit: int) -> list[Contact]:
        with self._scopes(PEOPLE_SCOPE) as ui:
            ui.cleanup.push(CLEAR_FIELD, self._source.clear)
            self._source.type(query)
            rows = self._waiter.until(
                lambda: typed_rows(self._source.rows()), SUGGESTION_WAIT_SECONDS
            )
        suggested = parse_suggestions(rows or [])
        recent = [
            contact for contact in reversed(self._recent.values()) if matches_query(contact, query)
        ]
        return unique_contacts([*suggested, *recent])[: max(limit, 0)]
