from __future__ import annotations

from collections.abc import Callable

from office_mcp.features.outlook.actions.gateway import MoveSheet, MoveTarget
from office_mcp.features.outlook.actions.move import require_first
from office_mcp.features.outlook.window import OutlookWindow, classify_window
from office_mcp.shared.ax import AxApp, as_bool, as_sequence
from office_mcp.shared.errors import (
    ActionNotVerifiedError,
    ActivationError,
    ElementNotFoundError,
    UiTimeoutError,
)
from office_mcp.shared.events import InputEvents, Key, Modifier, Shortcut
from office_mcp.shared.menus import MenuBar, PopupMenus
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize
from office_mcp.shared.tree import by_role_and_label, find, find_all
from office_mcp.shared.waiting import Waiter
from office_mcp.shared.windows import WindowLocator

MESSAGE_LIST = "Message List"
SEARCH_FIELD = "Search"
SEARCH_BOX = "search box"
MOVE_BUTTON = "Move"
CANCEL_BUTTON = "Cancel"
CHOOSE_FOLDER = ("Message", "Move", "Choose Folder...")
CHOOSE_FOLDER_SHORTCUT = Shortcut(Key.M, Modifier.COMMAND | Modifier.SHIFT)
TARGET_SEPARATOR = ", "
TABLE_DEPTH = 30
SHEET_DEPTH = 12
SHEET_WAIT = 5.0
FILTER_WAIT = 3.0


def parse_target(description: str) -> MoveTarget | None:
    folder, separator, account = description.rpartition(TARGET_SEPARATOR)
    if not separator or not folder.strip() or not account.strip():
        return None
    return MoveTarget(folder=folder.strip(), account=account.strip(), description=description)


class AxActions:
    def __init__(
        self,
        app: AxApp,
        windows: WindowLocator,
        menus: MenuBar,
        popups: PopupMenus,
        events: InputEvents,
        waiter: Waiter,
    ) -> None:
        self._app = app
        self._windows = windows
        self._menus = menus
        self._popups = popups
        self._events = events
        self._waiter = waiter

    def press_menu(self, path: tuple[str, ...], shortcut: Shortcut | None = None) -> None:
        self._menus.press(path, shortcut)

    def menu_titles(self, path: tuple[str, ...]) -> tuple[str, ...]:
        return tuple(
            entry.path[-1]
            for entry in self._menus.entries(path[0])
            if entry.path[:-1] == path and entry.enabled
        )

    def open_row_menu(self) -> Node:
        table = find(self._app, self._main(), self._is_message_list, TABLE_DEPTH, self._is_search)
        if table is None:
            raise ElementNotFoundError("the Outlook message list was not found")
        return self._popups.open(table, "AXShowMenu")

    def choose(self, menu: Node, path: tuple[str, ...]) -> None:
        self._popups.choose(menu, path)

    def checked(self, menu: Node, path: tuple[str, ...]) -> bool | None:
        return self._popups.checked(menu, path)

    def close_menu(self, menu: Node | None = None) -> None:
        self._popups.close(menu)

    def open_move_sheet(self) -> MoveSheet:
        self._menus.press(CHOOSE_FOLDER, CHOOSE_FOLDER_SHORTCUT)
        sheet = self._waiter.require(
            self._find_sheet,
            lambda: UiTimeoutError("the Move sheet did not open"),
            SHEET_WAIT,
        )
        return MoveSheet(sheet)

    def filter_move_sheet(self, sheet: MoveSheet, text: str) -> tuple[MoveTarget, ...]:
        box = self._require(sheet.handle, self._is_search_box, "the Move sheet search box")
        self._app.set_attribute(box, "AXValue", text)
        wanted = normalize(text)
        self._waiter.until(lambda: self._filtered(sheet, wanted), FILTER_WAIT)
        self._waiter.settle(lambda: len(self._groups(sheet.handle)))
        targets = (parse_target(self._description(group)) for group in self._groups(sheet.handle))
        return tuple(target for target in targets if target is not None)

    def pick_move_target(self, sheet: MoveSheet, target: MoveTarget) -> None:
        self._require_front_sheet(sheet)
        descriptions = [self._description(group) for group in self._groups(sheet.handle)]
        if descriptions.count(target.description) != 1:
            raise ActionNotVerifiedError(
                f"the Move sheet changed while picking {target.folder!r}; nothing was moved"
            )
        require_first(target, descriptions)
        box = self._require(sheet.handle, self._is_search_box, "the Move sheet search box")
        self._app.set_attribute(box, "AXFocused", True)
        pid = self._app.require_pid()
        for _ in descriptions:
            self._events.press(pid, Shortcut(Key.UP))

    def confirm_move(self, sheet: MoveSheet) -> None:
        button = self._require(sheet.handle, self._button(MOVE_BUTTON), "the Move button")
        self._app.perform(button, "AXPress", MOVE_BUTTON)

    def cancel_move(self, sheet: MoveSheet) -> None:
        button = self._require(sheet.handle, self._button(CANCEL_BUTTON), "the Cancel button")
        self._app.perform(button, "AXPress", CANCEL_BUTTON)

    def sheet_open(self, sheet: MoveSheet) -> bool:
        current = self._find_sheet()
        return current is not None and self._app.same(current, sheet.handle)

    def _main(self) -> object:
        self._app.require_trust()
        for window in self._app.windows():
            if classify_window(self._windows.probe(window)) is OutlookWindow.MAIN:
                return window
        raise ElementNotFoundError("the main Outlook mail window is not open")

    def _find_sheet(self) -> object | None:
        children: list[object] = list(as_sequence(self._app.attribute(self._main(), "AXChildren")))
        return next((child for child in children if self._is_move_sheet(child)), None)

    def _is_move_sheet(self, element: object) -> bool:
        if self._role(element) != "AXSheet":
            return False
        return find(self._app, element, self._is_search_box, SHEET_DEPTH) is not None

    def _require_front_sheet(self, sheet: MoveSheet) -> None:
        if as_bool(self._app.attribute(self._app.application(), "AXFrontmost")) is not True:
            raise ActivationError("Outlook is not the frontmost app, so no folder was clicked")
        if not self.sheet_open(sheet):
            raise ActionNotVerifiedError("the Move sheet is no longer open; nothing was moved")

    def _filtered(self, sheet: MoveSheet, wanted: str) -> bool:
        groups = self._groups(sheet.handle)
        return all(wanted in normalize(self._description(group)) for group in groups)

    def _groups(self, sheet: object) -> list[object]:
        area = find(self._app, sheet, self._has_role("AXScrollArea"), SHEET_DEPTH)
        if area is None:
            return []
        return find_all(self._app, area, self._is_target_group, SHEET_DEPTH)

    def _is_target_group(self, element: object) -> bool:
        return (
            self._role(element) == "AXGroup"
            and parse_target(self._description(element)) is not None
        )

    def _is_message_list(self, element: object) -> bool:
        return self._role(element) == "AXTable" and self._description(element) == MESSAGE_LIST

    def _is_search(self, element: object) -> bool:
        return self._role(element) == "AXTextField" and self._description(element) == SEARCH_FIELD

    def _is_search_box(self, element: object) -> bool:
        return self._role(element) == "AXTextField" and self._description(element) == SEARCH_BOX

    def _button(self, label: str) -> Callable[[object], bool]:
        return by_role_and_label(self._app, "AXButton", label)

    def _has_role(self, role: str) -> Callable[[object], bool]:
        return lambda element: self._role(element) == role

    def _require(self, root: object, matches: Callable[[object], bool], what: str) -> object:
        found = find(self._app, root, matches, SHEET_DEPTH)
        if found is None:
            raise ElementNotFoundError(f"{what} was not found")
        return found

    def _role(self, element: object) -> str:
        return self._app.text(element, "AXRole")

    def _description(self, element: object) -> str:
        return " ".join(self._app.text(element, "AXDescription").split())
