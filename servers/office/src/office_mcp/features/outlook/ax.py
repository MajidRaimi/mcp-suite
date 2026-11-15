from __future__ import annotations

from collections.abc import Callable

from office_mcp.features.outlook.rows import SearchScope, SearchState
from office_mcp.features.outlook.suggestions import close_suggestions
from office_mcp.features.outlook.window import (
    FocusedTab,
    OutlookMode,
    OutlookModule,
    OutlookWindow,
    account_from_title,
    classify_window,
    detect_mode,
)
from office_mcp.shared.ax import AxApp, as_bool, as_int, as_sequence, custom_action_label
from office_mcp.shared.errors import ElementNotFoundError, UiTimeoutError
from office_mcp.shared.events import InputEvents, Key, Shortcut
from office_mcp.shared.menus import PopupMenus, refuse_forbidden
from office_mcp.shared.node import Node
from office_mcp.shared.tree import by_role_and_label, element_label, find, find_all
from office_mcp.shared.windows import WindowLocator, WindowProbe

OUTLOOK_BUNDLE = "com.microsoft.Outlook"
OUTLOOK_APP_NAME = "Microsoft Outlook"
MESSAGE_LIST = "Message List"
MESSAGE_LIST_HEADER = "Message List Header"
HEADER_DETAILS = "Message Header Details"
READING_PANE = "Reading Pane"
PANE_CLIMB = 4
SEARCH_FIELD = "Search"
SEARCH_SCOPE = "Search Scope"
PERFORM_SEARCH = "Perform Search"
CANCEL_SEARCH = "Cancel Search"
SEARCH_STUCK = "Outlook did not leave search in time"
ESCAPE_ATTEMPTS = 3
LEAVE_SECONDS = 3.0
SEARCH_DEPTH = 30
FIELD_DEPTH = 4
POPOVER_DEPTH = 12
STATUS_SCAN_DEPTH = 4
STATUS_SCAN_SKIP = frozenset({"AXTable", "AXOutline"})
HEAVY_ROLES = frozenset({"AXTable", "AXOutline", "AXWebArea"})
FOCUSED_SELECTED = "Focused Inbox Selected"
OTHER_SELECTED = "Other Inbox Selected"
TAB_TOGGLE_MARK = "Inbox Selected"
GOT_IT = "Got it"
CUSTOM_ACTION_PREFIX = "Name:"
NO_MAIN_WINDOW = (
    "The main Outlook mail window is not open, so Outlook's state could not be checked. Open "
    "the main mail window in Outlook."
)


class AxOutlook:
    def __init__(
        self, app: AxApp, windows: WindowLocator, popups: PopupMenus, events: InputEvents
    ) -> None:
        self._app = app
        self._windows = windows
        self._popups = popups
        self._events = events
        self._main_handle: object | None = None

    def main_probe(self) -> WindowProbe:
        return self._windows.probe(self._main())

    def window_title(self) -> str:
        return self._app.text(self._main(), "AXTitle")

    def account(self) -> str | None:
        return account_from_title(self.window_title())

    def mode(self) -> OutlookMode:
        return detect_mode(self.main_probe())

    def shallow_button_labels(self, launch: bool = True) -> list[str]:
        labels: list[str] = []
        stack: list[tuple[object, int]] = [(self._main() if launch else self._existing_main(), 0)]
        while stack:
            element, depth = stack.pop()
            role = self._role(element)
            if role == "AXButton":
                label = self._app.text(element, "AXDescription")
                if label:
                    labels.append(label)
            if depth >= STATUS_SCAN_DEPTH or role in STATUS_SCAN_SKIP:
                continue
            children = as_sequence(self._app.attribute(element, "AXChildren"))
            stack.extend((child, depth + 1) for child in children)
        return labels

    def folder_rows(self) -> list[Node]:
        outline = self._find(self._main(), lambda e: self._role(e) == "AXOutline")
        return [self._folder_row(row) for row in self._rows(outline)] if outline else []

    def message_rows(self) -> list[Node]:
        table = self._find(self._main(), self._is_message_list)
        return [self._message_row(row) for row in self._rows(table)] if table else []

    def select(self, node: Node) -> None:
        self._app.set_attribute(self._handle(node), "AXSelected", True)

    def deselect(self, node: Node) -> None:
        self._app.set_attribute(self._handle(node), "AXSelected", False)

    def collapse_selected(self) -> None:
        table = self._find(self._main(), self._is_message_list)
        if table is None:
            raise ElementNotFoundError("the Outlook message list was not found")
        self._app.set_attribute(table, "AXFocused", True)
        self._events.press(self._app.require_pid(), Shortcut(Key.LEFT))

    def press(self, node: Node) -> None:
        handle = self._handle(node)
        cells = list(as_sequence(self._app.attribute(handle, "AXChildren")))
        self._app.perform(cells[0] if cells else handle, "AXPress", node.label())

    def reading_pane(self) -> Node | None:
        header = self._find(self._main(), lambda e: self._description(e) == HEADER_DETAILS)
        if header is None:
            return None
        container = self._app.attribute(header, "AXParent")
        if container is None:
            return None
        top = self._app.snapshot(container, prune=lambda n: n.description == MESSAGE_LIST)
        web = self._reading_web_area(container)
        body = (self._app.snapshot(web),) if web is not None else ()
        return Node(role="AXGroup", children=(top, *body))

    def toolbar_button(self, name: str) -> Node | None:
        button = self._find(
            self._main(), lambda e: self._role(e) == "AXButton" and self._description(e) == name
        )
        return Node(role="AXButton", description=name, handle=button) if button else None

    def press_button(self, node: Node) -> None:
        self._app.perform(self._handle(node), "AXPress", node.label())

    def focused_tab(self) -> FocusedTab | None:
        toggle = self._tab_toggle()
        return tab_state(self._app, toggle) if toggle is not None else None

    def switch_tab(self, tab: FocusedTab) -> None:
        if self.focused_tab() == tab:
            return
        toggle = self._tab_toggle()
        if toggle is None:
            raise ElementNotFoundError("the Focused and Other switch is not shown in this folder")
        self._app.perform(toggle, "AXPress", "Focused and Other switch")
        self._app.waiter.require(
            lambda: self.focused_tab() == tab,
            lambda: UiTimeoutError(f"Outlook did not switch to the {tab.value} tab in time"),
        )

    def search_state(self) -> SearchState | None:
        if self.mode() is not OutlookMode.SEARCH:
            return None
        field = self._search_field()
        if field is None:
            return None
        query = self._app.text(field, "AXValue")
        return SearchState(query=query, scope=self._current_scope(field))

    def run_search(self, query: str, scope: SearchScope) -> None:
        self.cancel_search()
        field = self._search_field()
        if field is None:
            raise ElementNotFoundError("the Outlook search field was not found")
        self._app.set_attribute(field, "AXFocused", True)
        self._app.set_attribute(field, "AXValue", query)
        if self._current_scope(field) is not scope:
            self._choose_scope(field, scope)
        if self.mode() is not OutlookMode.SEARCH:
            self._perform_search(field)
        self._app.waiter.require(
            lambda: self.mode() is OutlookMode.SEARCH,
            lambda: UiTimeoutError("Outlook did not show search results in time"),
        )

    def _perform_search(self, field: object) -> None:
        if as_bool(self._app.attribute(field, "AXFocused")) is not True:
            self._app.set_attribute(field, "AXFocused", True)
        perform = self._app.waiter.require(
            lambda: self._find_in(field, PERFORM_SEARCH),
            lambda: ElementNotFoundError(f"the Outlook {PERFORM_SEARCH!r} button was not found"),
        )
        self._app.perform(perform, "AXPress", PERFORM_SEARCH)

    def cancel_search(self) -> None:
        if self.mode() is not OutlookMode.SEARCH:
            return
        button = find(
            self._app,
            self._main(),
            by_role_and_label(self._app, "AXButton", CANCEL_SEARCH),
            FIELD_DEPTH * 2,
            prune=self._is_heavy,
        )
        if button is not None:
            self._app.perform(button, "AXPress", CANCEL_SEARCH)
            self._app.waiter.require(self._left_search, lambda: UiTimeoutError(SEARCH_STUCK))
            return
        for _ in range(ESCAPE_ATTEMPTS):
            self.close_search_suggestions()
            self._escape_search()
            if self._app.waiter.until(self._left_search, LEAVE_SECONDS):
                return
        raise UiTimeoutError(SEARCH_STUCK)

    def _left_search(self) -> bool:
        return self.mode() is not OutlookMode.SEARCH

    def _escape_search(self) -> None:
        field = self._search_field()
        if field is None:
            raise ElementNotFoundError("the Outlook search field was not found to leave the search")
        self._app.set_attribute(field, "AXFocused", True)
        self._events.press(self._app.require_pid(), Shortcut(Key.ESCAPE))

    def show_module(self, module: OutlookModule) -> None:
        button = self._find(self._main(), lambda e: self._is_module_button(e, module))
        if button is None:
            raise ElementNotFoundError(f"the {module.value} module button was not found")
        self._app.perform(button, "AXPress", module.value)
        wanted_calendar = module is OutlookModule.CALENDAR
        self._app.waiter.require(
            lambda: (self.mode() is OutlookMode.CALENDAR) == wanted_calendar,
            lambda: UiTimeoutError(f"Outlook did not switch to {module.value} in time"),
        )

    def row_actions(self, row: Node) -> tuple[str, ...]:
        raw_names = self._app.actions(self._handle(row))
        return tuple(
            custom_action_label(raw) for raw in raw_names if raw.startswith(CUSTOM_ACTION_PREFIX)
        )

    def perform_row_action(self, row: Node, label: str) -> None:
        refuse_forbidden((label,))
        handle = self._handle(row)
        raw = self._app.custom_action(handle, label)
        if raw is None:
            raise ElementNotFoundError(f"the conversation row has no {label!r} action")
        self._app.perform(handle, raw, label)

    def alert_sheets(self) -> list[Node]:
        return [
            self._app.snapshot(child)
            for window in self._app.windows()
            for child in as_sequence(self._app.attribute(window, "AXChildren"))
            if self._role(child) == "AXSheet"
        ]

    def close_search_suggestions(self) -> int:
        return close_suggestions(self._app, self._events, self._app.waiter)

    def dismiss_known_popovers(self) -> int:
        dismissed = 0
        for window in self._app.windows():
            for popover in find_all(
                self._app, window, lambda e: self._role(e) == "AXPopover", POPOVER_DEPTH
            ):
                button = find(self._app, popover, by_role_and_label(self._app, "AXButton", GOT_IT))
                if button is not None:
                    self._app.perform(button, "AXPress", GOT_IT)
                    dismissed += 1
        return dismissed

    def _main(self) -> object:
        self._app.require_trust()
        cached = self._cached_main()
        if cached is not None:
            return cached
        window = self._app.window_or_launch(self._is_main)
        self._main_handle = window
        return window

    def _existing_main(self) -> object:
        self._app.require_trust()
        cached = self._cached_main()
        if cached is not None:
            return cached
        window = next((found for found in self._app.windows() if self._is_main(found)), None)
        if window is None:
            raise ElementNotFoundError(NO_MAIN_WINDOW)
        self._main_handle = window
        return window

    def _cached_main(self) -> object | None:
        handle = self._main_handle
        if handle is not None and any(self._app.same(handle, w) for w in self._app.windows()):
            return handle
        self._main_handle = None
        return None

    def _is_main(self, window: object) -> bool:
        return classify_window(self._windows.probe(window)) is OutlookWindow.MAIN

    def _is_message_list(self, element: object) -> bool:
        return self._role(element) == "AXTable" and self._description(element) == MESSAGE_LIST

    def _is_module_button(self, element: object, module: OutlookModule) -> bool:
        return self._role(element) == "AXButton" and self._app.text(element, "AXHelp") == (
            module.value
        )

    def _tab_toggle(self) -> object | None:
        header = self._find(self._main(), lambda e: self._description(e) == MESSAGE_LIST_HEADER)
        if header is None:
            return None
        return find(self._app, header, self._is_tab_toggle, FIELD_DEPTH)

    def _is_tab_toggle(self, element: object) -> bool:
        return self._role(element) == "AXCheckBox" and TAB_TOGGLE_MARK in element_label(
            self._app, element
        )

    def _search_field(self) -> object | None:
        return find(
            self._app,
            self._main(),
            lambda e: self._role(e) == "AXTextField" and self._description(e) == SEARCH_FIELD,
            FIELD_DEPTH,
            prune=self._is_heavy,
        )

    def _current_scope(self, field: object) -> SearchScope:
        button = self._find_in(field, SEARCH_SCOPE, "AXPopUpButton")
        value = self._app.text(button, "AXValue") if button is not None else ""
        return parse_scope(value)

    def _choose_scope(self, field: object, scope: SearchScope) -> None:
        button = self._find_in(field, SEARCH_SCOPE, "AXPopUpButton")
        if button is None:
            raise ElementNotFoundError("the Outlook search scope menu was not found")
        menu = self._popups.open(button, "AXPress")
        self._popups.choose(menu, (scope.value,))
        self._app.waiter.require(
            lambda: self._current_scope(field) is scope,
            lambda: UiTimeoutError(f"Outlook did not switch the search scope to {scope.value}"),
        )

    def _find_in(self, root: object, label: str, role: str = "AXButton") -> object | None:
        return find(self._app, root, by_role_and_label(self._app, role, label), FIELD_DEPTH)

    def _require_in(self, root: object, label: str) -> object:
        found = self._find_in(root, label)
        if found is None:
            raise ElementNotFoundError(f"the Outlook {label!r} button was not found")
        return found

    def _reading_web_area(self, start: object) -> object | None:
        container: object | None = start
        for _ in range(PANE_CLIMB):
            if container is None:
                return None
            web = self._find(container, self._is_reading_pane)
            if web is not None:
                return web
            container = self._app.attribute(container, "AXParent")
        return None

    def _is_reading_pane(self, element: object) -> bool:
        return self._role(element) == "AXWebArea" and self._description(element) == READING_PANE

    def _rows(self, container: object) -> list[object]:
        return list(as_sequence(self._app.attribute(container, "AXRows")))

    def _message_row(self, row: object) -> Node:
        return Node(
            role="AXRow",
            description=self._first_cell_description(row),
            handle=row,
            selected=as_bool(self._app.attribute(row, "AXSelected")) is True,
        )

    def _folder_row(self, row: object) -> Node:
        return Node(
            role="AXRow",
            description=self._first_cell_description(row),
            handle=row,
            level=as_int(self._app.attribute(row, "AXDisclosureLevel")),
            selected=as_bool(self._app.attribute(row, "AXSelected")) is True,
        )

    def _first_cell_description(self, row: object) -> str:
        cells = list(as_sequence(self._app.attribute(row, "AXChildren")))
        return self._description(cells[0]) if cells else ""

    def _role(self, element: object) -> str:
        return self._app.text(element, "AXRole")

    def _description(self, element: object) -> str:
        return self._app.text(element, "AXDescription")

    def _is_heavy(self, element: object) -> bool:
        return self._role(element) in HEAVY_ROLES

    def _is_search_field(self, element: object) -> bool:
        return self._role(element) == "AXTextField" and self._description(element) == SEARCH_FIELD

    def _find(
        self, root: object, matches: Callable[[object], bool], depth: int = SEARCH_DEPTH
    ) -> object | None:
        return find(self._app, root, matches, depth, prune=self._is_search_field)

    @staticmethod
    def _handle(node: Node) -> object:
        if node.handle is None:
            raise ElementNotFoundError("cannot act on an element without a handle")
        return node.handle


def tab_state(app: AxApp, toggle: object) -> FocusedTab | None:
    value = as_int(app.attribute(toggle, "AXValue"))
    if value == 1:
        return FocusedTab.OTHER
    if value == 0:
        return FocusedTab.FOCUSED
    label = element_label(app, toggle)
    if label.startswith(OTHER_SELECTED):
        return FocusedTab.OTHER
    if label.startswith(FOCUSED_SELECTED):
        return FocusedTab.FOCUSED
    return None


def parse_scope(value: str) -> SearchScope:
    try:
        return SearchScope(value.strip())
    except ValueError:
        return SearchScope.MAILBOX
