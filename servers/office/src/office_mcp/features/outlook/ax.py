from __future__ import annotations

from collections.abc import Callable
from typing import Any

from office_mcp.shared.ax import AxApp, as_sequence, as_text
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.node import Node

OUTLOOK_BUNDLE = "com.microsoft.Outlook"
OUTLOOK_APP_NAME = "Microsoft Outlook"
MAIN_WINDOW_MARK = "•"
MESSAGE_LIST = "Message List"
HEADER_DETAILS = "Message Header Details"
READING_PANE = "Reading Pane"
PANE_CLIMB = 4
SEARCH_FIELD = "Search"
SEARCH_DEPTH = 30


class AxOutlook:
    def __init__(self, app: AxApp | None = None) -> None:
        self._app = app or AxApp(OUTLOOK_BUNDLE, OUTLOOK_APP_NAME)

    def trusted(self) -> bool:
        return self._app.trusted()

    def running(self) -> bool:
        return self._app.running()

    def window_title(self) -> str:
        return as_text(self._app.attribute(self._main(), "AXTitle"))

    def folder_rows(self) -> list[Node]:
        outline = self._find(self._main(), lambda e: self._role(e) == "AXOutline")
        return [self._row(row) for row in self._rows(outline)] if outline is not None else []

    def message_rows(self) -> list[Node]:
        table = self._find(
            self._main(),
            lambda e: (
                self._role(e) == "AXTable"
                and as_text(self._app.attribute(e, "AXDescription")) == MESSAGE_LIST
            ),
        )
        return [self._row(row) for row in self._rows(table)] if table is not None else []

    def select(self, node: Node) -> None:
        self._app.set_attribute(self._handle(node), "AXSelected", True)

    def press(self, node: Node) -> None:
        handle = self._handle(node)
        cells = list(as_sequence(self._app.attribute(handle, "AXChildren")))
        self._app.perform(cells[0] if cells else handle, "AXPress", node.label())

    def reading_pane(self) -> Node | None:
        header = self._find(
            self._main(),
            lambda e: as_text(self._app.attribute(e, "AXDescription")) == HEADER_DETAILS,
        )
        if header is None:
            return None
        container = self._app.attribute(header, "AXParent")
        top = self._app.snapshot(container, skip=lambda n: n.description == MESSAGE_LIST)
        web = self._reading_web_area(container)
        body = (self._app.snapshot(web),) if web is not None else ()
        return Node(role="AXGroup", children=(top, *body))

    def _reading_web_area(self, start: Any) -> Any:
        container = start
        for _ in range(PANE_CLIMB):
            web = self._find(container, self._is_reading_pane)
            if web is not None:
                return web
            container = self._app.attribute(container, "AXParent")
            if container is None:
                return None
        return None

    def _is_reading_pane(self, element: Any) -> bool:
        return (
            self._role(element) == "AXWebArea"
            and as_text(self._app.attribute(element, "AXDescription")) == READING_PANE
        )

    def toolbar_button(self, name: str) -> Node | None:
        button = self._find(
            self._main(),
            lambda e: (
                self._role(e) == "AXButton"
                and as_text(self._app.attribute(e, "AXDescription")) == name
            ),
        )
        return Node(role="AXButton", description=name, handle=button) if button else None

    def press_button(self, node: Node) -> None:
        self._app.perform(self._handle(node), "AXPress", node.label())

    def _main(self) -> Any:
        self._app.require_trust()
        return self._app.window(lambda title: MAIN_WINDOW_MARK in title)

    def _rows(self, container: Any) -> list[Any]:
        return list(as_sequence(self._app.attribute(container, "AXRows")))

    def _row(self, row: Any) -> Node:
        cells = list(as_sequence(self._app.attribute(row, "AXChildren")))
        description = as_text(self._app.attribute(cells[0], "AXDescription")) if cells else ""
        return Node(role="AXRow", description=description, handle=row)

    def _role(self, element: Any) -> str:
        return as_text(self._app.attribute(element, "AXRole"))

    def _find(self, root: Any, matches: Callable[[Any], bool]) -> Any:
        stack: list[tuple[Any, int]] = [(root, 0)]
        seen: set[int] = set()
        while stack:
            element, depth = stack.pop(0)
            if hash(element) in seen:
                continue
            seen.add(hash(element))
            if matches(element):
                return element
            if depth >= SEARCH_DEPTH or self._is_search_field(element):
                continue
            children = as_sequence(self._app.attribute(element, "AXChildren"))
            stack.extend((child, depth + 1) for child in children)
        return None

    def _is_search_field(self, element: Any) -> bool:
        return (
            self._role(element) == "AXTextField"
            and as_text(self._app.attribute(element, "AXDescription")) == SEARCH_FIELD
        )

    @staticmethod
    def _handle(node: Node) -> Any:
        if node.handle is None:
            raise OfficeError("cannot act on an element without a handle")
        return node.handle
