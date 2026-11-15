from __future__ import annotations

from collections.abc import Sequence
from contextlib import suppress
from dataclasses import dataclass

from office_mcp.shared.ax import MENU_OPEN_CODES, AxApp, as_bool, as_int, as_sequence
from office_mcp.shared.errors import (
    ActionNotVerifiedError,
    ElementNotFoundError,
    ForbiddenActionError,
    MenuItemDisabledError,
    MenuStuckError,
    OfficeError,
)
from office_mcp.shared.events import (
    InputEvents,
    Key,
    Shortcut,
    decode_menu_modifiers,
    menu_key,
)
from office_mcp.shared.geometry import is_empty_size
from office_mcp.shared.node import Node
from office_mcp.shared.tree import find
from office_mcp.shared.waiting import Waiter

MENU_ROLE = "AXMenu"
MENU_BAR_ROLE = "AXMenuBar"
APPLICATION_ROLE = "AXApplication"
MENU_ITEM_ROLE = "AXMenuItem"
CELL_ROLE = "AXCell"
ROW_ROLE = "AXRow"
CHOOSE_ROW = Shortcut(Key.RETURN)
MENU_WINDOW_TITLE = "menuWindow"
ITEM_ROLES = frozenset(
    {
        "AXMenuItem",
        "AXMenuBarItem",
        "AXCell",
        "AXCheckBox",
        "AXRadioButton",
        "AXButton",
        "AXMenuButton",
    }
)
ELEMENT_SEARCH_DEPTH = 6
PARENT_LEVELS = 4
WINDOW_SEARCH_DEPTH = 40
CLOSE_WAIT = 1.0
OPEN_WAIT = 3.0
LABEL_DEPTH = 3
CHECKED_VALUES = frozenset({"1", "true"})
FORBIDDEN_MENU_WORDS: frozenset[str] = frozenset(
    {
        "Permanently Delete",
        "Empty Folder",
        "Delete Folder",
        "Sign Out",
        "Recall",
        "Resend",
        "Ignore Conversation",
        "Report Phishing",
        "Report Junk",
        "Report this message",
        "Block",
        "Unsubscribe",
        "Always Move to Other",
        "Always Move to Focused",
        "Retention",
        "Assign policy",
        "Mark All as Read",
        "Work Offline",
        "Turn Off Reminders",
        "Clear Application Data",
        "Revert to Legacy Outlook",
        "Legacy Outlook",
        "Quit Outlook",
        "Join",
        "Meet now",
        "Pin for everyone",
        "Toggle Unread",
    }
)
FOLDED_FORBIDDEN = frozenset(" ".join(word.casefold().split()) for word in FORBIDDEN_MENU_WORDS)


@dataclass(frozen=True)
class MenuEntry:
    path: tuple[str, ...]
    enabled: bool
    checked: bool
    shortcut: Shortcut | None
    handle: object


def fold(text: str) -> str:
    return " ".join(text.replace("…", "...").casefold().split())


def is_forbidden(path: Sequence[str]) -> bool:
    return any(fold(element) in FOLDED_FORBIDDEN for element in path)


def show_path(path: Sequence[str]) -> str:
    return " > ".join(path)


def refuse_forbidden(path: Sequence[str]) -> None:
    if is_forbidden(path):
        raise ForbiddenActionError(f"the menu item {show_path(path)} is never used by office-mcp")


def disabled_error(path: Sequence[str], app_name: str) -> MenuItemDisabledError:
    return MenuItemDisabledError(
        f"the menu item {show_path(path)} is disabled; {app_name} must be the frontmost app "
        "with the right item selected"
    )


def item_matches(node: Node, title: str) -> bool:
    return node.role in ITEM_ROLES and fold(node.label()) == fold(title)


def find_item(menu: Node, path: tuple[str, ...]) -> Node | None:
    found: Node | None = menu
    for title in path:
        found = child_item(found, title) if found is not None else None
    return found if path else None


def child_item(scope: Node, title: str) -> Node | None:
    return next(
        (node for node in scope.walk() if node is not scope and item_matches(node, title)), None
    )


class MenuBar:
    def __init__(self, app: AxApp) -> None:
        self._app = app

    def entries(self, top: str) -> list[MenuEntry]:
        menu = self._top_menu(top)
        return self._collect(menu, (top,)) if menu is not None else []

    def entry(self, path: tuple[str, ...], shortcut: Shortcut | None = None) -> MenuEntry:
        if len(path) < 2:
            raise ElementNotFoundError(f"menu path {show_path(path)} needs a menu and an item")
        menu = self._top_menu(path[0])
        if menu is None:
            raise ElementNotFoundError(f"{self._app.app_name} has no {path[0]} menu")
        found: MenuEntry | None = None
        for depth in range(1, len(path)):
            leaf = depth == len(path) - 1
            candidates = [
                item for item in self._items(menu) if self._app.text(item, "AXTitle") == path[depth]
            ]
            chosen = self._choose(candidates, path[: depth + 1], shortcut if leaf else None)
            found = self._entry(chosen, path[: depth + 1])
            if not leaf:
                submenu = self._submenu(chosen)
                if submenu is None:
                    raise ElementNotFoundError(f"{show_path(path[: depth + 1])} has no submenu")
                menu = submenu
        if found is None:
            raise ElementNotFoundError(f"menu item {show_path(path)} was not found")
        return found

    def press(self, path: tuple[str, ...], shortcut: Shortcut | None = None) -> None:
        refuse_forbidden(path)
        found = self.entry(path, shortcut)
        if not found.enabled:
            raise disabled_error(path, self._app.app_name)
        if self._app.text(found.handle, "AXTitle") != path[-1]:
            raise ElementNotFoundError(f"the menu item {show_path(path)} changed before pressing")
        self._app.perform(found.handle, "AXPress", show_path(path))

    def _top_menu(self, top: str) -> object | None:
        bar = self._app.attribute(self._app.application(), "AXMenuBar")
        if bar is None:
            return None
        for item in as_sequence(self._app.attribute(bar, "AXChildren")):
            if self._app.text(item, "AXTitle") == top:
                return self._submenu(item)
        return None

    def _submenu(self, item: object) -> object | None:
        return next(
            (
                child
                for child in as_sequence(self._app.attribute(item, "AXChildren"))
                if self._app.text(child, "AXRole") == MENU_ROLE
            ),
            None,
        )

    def _items(self, menu: object) -> list[object]:
        return [
            child
            for child in as_sequence(self._app.attribute(menu, "AXChildren"))
            if self._app.text(child, "AXRole") == MENU_ITEM_ROLE
        ]

    def _choose(
        self, candidates: list[object], path: tuple[str, ...], shortcut: Shortcut | None
    ) -> object:
        if shortcut is not None:
            candidates = [item for item in candidates if self._shortcut(item) == shortcut]
        if not candidates:
            raise ElementNotFoundError(f"menu item {show_path(path)} was not found")
        if len(candidates) > 1:
            raise ElementNotFoundError(
                f"menu item {show_path(path)} matches {len(candidates)} items; "
                "pass its shortcut to pick one"
            )
        return candidates[0]

    def _entry(self, item: object, path: tuple[str, ...]) -> MenuEntry:
        return MenuEntry(
            path=path,
            enabled=as_bool(self._app.attribute(item, "AXEnabled")) is not False,
            checked=bool(self._app.text(item, "AXMenuItemMarkChar").strip()),
            shortcut=self._shortcut(item),
            handle=item,
        )

    def _collect(self, menu: object, prefix: tuple[str, ...]) -> list[MenuEntry]:
        collected: list[MenuEntry] = []
        for item in self._items(menu):
            title = self._app.text(item, "AXTitle")
            if not title:
                continue
            path = (*prefix, title)
            collected.append(self._entry(item, path))
            submenu = self._submenu(item)
            if submenu is not None:
                collected.extend(self._collect(submenu, path))
        return collected

    def _shortcut(self, item: object) -> Shortcut | None:
        key = menu_key(
            self._app.text(item, "AXMenuItemCmdChar"),
            as_int(self._app.attribute(item, "AXMenuItemCmdVirtualKey")),
        )
        if key is None:
            return None
        modifiers = as_int(self._app.attribute(item, "AXMenuItemCmdModifiers")) or 0
        return Shortcut(key, decode_menu_modifiers(modifiers))


class PopupMenus:
    def __init__(self, app: AxApp, events: InputEvents, waiter: Waiter) -> None:
        self._app = app
        self._events = events
        self._waiter = waiter

    def open(self, element: object, action: str = "AXShowMenu") -> Node:
        self._app.perform(element, action, "menu", tolerate=MENU_OPEN_CODES)
        menu = self._waiter.require(
            lambda: self._locate(element),
            lambda: ElementNotFoundError(f"no menu opened in {self._app.app_name}"),
            OPEN_WAIT,
        )
        return self._app.snapshot(menu)

    def choose(self, menu: Node, path: tuple[str, ...]) -> None:
        refuse_forbidden(path)
        item = find_item(menu, path)
        if item is None and len(path) > 1:
            item = self._open_submenu(menu, path)
        if item is None:
            raise ElementNotFoundError(f"menu item {show_path(path)} was not found")
        self._press_item(item, path)

    def checked(self, menu: Node, path: tuple[str, ...]) -> bool | None:
        item = find_item(menu, path)
        if item is None:
            return None
        if item.handle is not None and self._app.text(item.handle, "AXMenuItemMarkChar").strip():
            return True
        return item.value.strip().casefold() in CHECKED_VALUES

    def close(self, menu: Node | None = None) -> None:
        targets = [menu.handle] if menu is not None and menu.handle is not None else []
        if not targets:
            targets = self._open_menus()
        for target in targets:
            self._close_one(target)

    def _open_submenu(self, menu: Node, path: tuple[str, ...]) -> Node | None:
        parent = find_item(menu, path[:-1])
        if parent is None or parent.handle is None:
            return None
        self._app.perform(parent.handle, "AXPress", show_path(path[:-1]), tolerate=MENU_OPEN_CODES)
        return self._waiter.until(lambda: find_item(self._refresh(menu), path), OPEN_WAIT)

    def _refresh(self, menu: Node) -> Node:
        return self._app.snapshot(menu.handle) if menu.handle is not None else menu

    def _press_item(self, item: Node, path: tuple[str, ...]) -> None:
        if item.enabled is False:
            raise disabled_error(path, self._app.app_name)
        if item.handle is None:
            raise ElementNotFoundError(f"menu item {show_path(path)} has no element")
        current = self._app.snapshot(item.handle, max_depth=LABEL_DEPTH)
        if fold(current.label()) != fold(path[-1]):
            raise ElementNotFoundError(f"the menu item {show_path(path)} changed before pressing")
        if item.role == CELL_ROLE:
            self._choose_row(item.handle, path)
            return
        self._app.perform(item.handle, "AXPress", show_path(path))

    def _choose_row(self, cell: object, path: tuple[str, ...]) -> None:
        row = self._app.attribute(cell, "AXParent")
        if row is None or self._app.text(row, "AXRole") != ROW_ROLE:
            raise ElementNotFoundError(f"the menu row for {show_path(path)} was not found")
        self._app.set_attribute(row, "AXSelected", True)
        if as_bool(self._app.attribute(row, "AXSelected")) is not True:
            raise ActionNotVerifiedError(f"the menu row {show_path(path)} could not be selected")
        self._events.press(self._app.require_pid(), CHOOSE_ROW)

    def _locate(self, element: object) -> object | None:
        return (
            self._in_element(element)
            or self._in_parents(element)
            or self._in_windows()
            or self._menu_window()
        )

    def _is_menu(self, element: object) -> bool:
        return self._app.text(element, "AXRole") == MENU_ROLE

    def _is_menu_bar(self, element: object) -> bool:
        return self._app.text(element, "AXRole") == MENU_BAR_ROLE

    def _in_element(self, element: object) -> object | None:
        return find(self._app, element, self._is_menu, ELEMENT_SEARCH_DEPTH)

    def _in_parents(self, element: object) -> object | None:
        current = element
        for _ in range(PARENT_LEVELS):
            parent = self._app.attribute(current, "AXParent")
            if parent is None or self._app.text(parent, "AXRole") == APPLICATION_ROLE:
                return None
            found = find(
                self._app, parent, self._is_menu, ELEMENT_SEARCH_DEPTH, prune=self._is_menu_bar
            )
            if found is not None:
                return found
            current = parent
        return None

    def _in_windows(self) -> object | None:
        for root in self._roots():
            found = find(
                self._app,
                root,
                self._is_menu,
                WINDOW_SEARCH_DEPTH,
                prune=self._is_menu_bar,
            )
            if found is not None:
                return found
        return None

    def _menu_window(self) -> object | None:
        return next(
            (
                window
                for window in self._app.windows()
                if self._app.text(window, "AXTitle") == MENU_WINDOW_TITLE
            ),
            None,
        )

    def _roots(self) -> list[object]:
        application = self._app.application()
        return [
            child
            for child in as_sequence(self._app.attribute(application, "AXChildren"))
            if self._app.text(child, "AXRole") != MENU_BAR_ROLE
        ]

    def _open_menus(self) -> list[object]:
        menus: list[object] = []
        found = self._in_windows()
        if found is not None:
            menus.append(found)
        window = self._menu_window()
        if window is not None:
            menus.append(window)
        return menus

    def _is_open(self, element: object) -> bool:
        if self._app.attribute(element, "AXRole") is None:
            return False
        return not is_empty_size(self._app.attribute(element, "AXSize"))

    def _close_one(self, target: object) -> None:
        if not self._is_open(target):
            return
        with suppress(OfficeError):
            self._app.perform(target, "AXCancel", "menu")
        if self._closed(target):
            return
        self._escape(target)
        if self._closed(target):
            return
        raise MenuStuckError(f"a menu in {self._app.app_name} did not close")

    def _escape(self, target: object) -> None:
        item = find(
            self._app,
            target,
            lambda element: self._app.text(element, "AXRole") in ITEM_ROLES,
            ELEMENT_SEARCH_DEPTH,
        )
        with suppress(OfficeError):
            self._app.set_attribute(item if item is not None else target, "AXFocused", True)
        self._events.press(self._app.require_pid(), Shortcut(Key.ESCAPE))

    def _closed(self, target: object) -> bool:
        return bool(self._waiter.until(lambda: not self._is_open(target), CLOSE_WAIT))
