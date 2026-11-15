from __future__ import annotations

from collections.abc import Callable

from office_mcp.shared.errors import ElementNotFoundError
from office_mcp.shared.events import Shortcut
from office_mcp.shared.node import Node


class FakeTeamsGateway:
    def __init__(self, screens: list[Node], on_press: Callable[[Node], None] | None = None) -> None:
        self.screens = screens
        self.on_press = on_press
        self.menus: dict[object, Node] = {}
        self.pressed: list[object] = []
        self.scrolled: list[str] = []
        self.focused: list[object] = []
        self.values: list[tuple[object, str]] = []
        self.keys: list[Shortcut] = []
        self.popups_closed = 0
        self.close_calls = 0
        self.snapshots = 0

    def snapshot(self) -> Node:
        self.snapshots += 1
        return self.screens[0] if len(self.screens) == 1 else self.screens.pop(0)

    def show(self, *screens: Node) -> None:
        self.screens[:] = list(screens)

    def press(self, node: Node) -> None:
        self.pressed.append(node.handle)
        if self.on_press is not None:
            self.on_press(node)

    def scroll_into_view(self, node: Node) -> None:
        self.scrolled.append(node.dom_id)

    def focus(self, node: Node) -> None:
        self.focused.append(node.handle)

    def set_value(self, node: Node, text: str) -> None:
        self.values.append((node.handle, text))

    def show_menu(self, node: Node) -> Node:
        menu = self.menus.get(node.handle)
        if menu is None:
            raise ElementNotFoundError(f"no menu scripted for {node.handle!r}")
        return menu

    def close_popups(self) -> int:
        self.close_calls += 1
        closed, self.popups_closed = self.popups_closed, 0
        return closed

    def press_key(self, shortcut: Shortcut) -> None:
        self.keys.append(shortcut)

    def element(self, node: Node) -> object:
        if node.handle is None:
            raise ElementNotFoundError("no handle")
        return node.handle
