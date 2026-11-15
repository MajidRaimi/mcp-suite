from __future__ import annotations

from collections.abc import Iterator
from contextlib import suppress

from office_mcp.shared.ax import MENU_OPEN_CODES, AxApp
from office_mcp.shared.errors import ElementNotFoundError, MenuStuckError, OfficeError
from office_mcp.shared.events import InputEvents, Key, Shortcut
from office_mcp.shared.menus import ITEM_ROLES
from office_mcp.shared.node import Node
from office_mcp.shared.waiting import Waiter

TEAMS_BUNDLE = "com.microsoft.teams2"
TEAMS_APP_NAME = "Microsoft Teams"
STANDARD_WINDOW = "AXStandardWindow"
MENU_ROLE = "AXMenu"
DIALOG = "AXApplicationDialog"
SETTLE_TIMEOUT = 12.0
SETTLED_MIN_NODES = 200
MENU_WAIT = 3.0
CLOSE_WAIT = 1.5
MAX_POPUP_ROUNDS = 5
ESCAPE = Shortcut(Key.ESCAPE)
PERSISTENT_PANEL_PREFIXES = ("calv2-sf-tracking-view",)


def is_popup(node: Node) -> bool:
    return node.role in (MENU_ROLE, DIALOG) or node.subrole == DIALOG


def outside_panels(node: Node) -> Iterator[Node]:
    if node.dom_id.startswith(PERSISTENT_PANEL_PREFIXES):
        return
    yield node
    for child in node.children:
        yield from outside_panels(child)


def find_popup(root: Node) -> Node | None:
    return next((node for node in outside_panels(root) if is_popup(node)), None)


def find_menu(root: Node) -> Node | None:
    return root.find(lambda node: node.role == MENU_ROLE)


class AxTeams:
    def __init__(self, app: AxApp, events: InputEvents, waiter: Waiter) -> None:
        self._app = app
        self._events = events
        self._waiter = waiter
        self._settled = False

    def snapshot(self) -> Node:
        self._app.require_trust()
        if self._settled:
            return self._app.snapshot(self._window())
        node = self._settle()
        self._settled = True
        return node

    def press(self, node: Node) -> None:
        self._app.perform(self.element(node), "AXPress", node.label())

    def scroll_into_view(self, node: Node) -> None:
        self._app.perform(self.element(node), "AXScrollToVisible", node.label())

    def focus(self, node: Node) -> None:
        self._app.set_attribute(self.element(node), "AXFocused", True)

    def set_value(self, node: Node, text: str) -> None:
        self._app.set_attribute(self.element(node), "AXValue", text)

    def show_menu(self, node: Node) -> Node:
        self._app.perform(self.element(node), "AXShowMenu", node.label(), tolerate=MENU_OPEN_CODES)
        return self._waiter.require(
            lambda: find_menu(self.snapshot()),
            lambda: ElementNotFoundError(f"no menu opened for {node.label()[:60]!r} in Teams"),
            MENU_WAIT,
        )

    def close_popups(self) -> int:
        closed = 0
        for _ in range(MAX_POPUP_ROUNDS):
            popup = find_popup(self.snapshot())
            if popup is None:
                return closed
            self._escape(popup)
            if self._wait_closed(popup):
                closed += 1
        if find_popup(self.snapshot()) is not None:
            raise MenuStuckError("a Teams menu or dialog did not close")
        return closed

    def press_key(self, shortcut: Shortcut) -> None:
        self._events.press(self._app.require_pid(), shortcut)

    def element(self, node: Node) -> object:
        if node.handle is None:
            raise ElementNotFoundError("cannot act on a Teams element without a handle")
        return node.handle

    def _escape(self, popup: Node) -> None:
        target = popup.find(lambda node: node is not popup and node.role in ITEM_ROLES) or popup
        with suppress(OfficeError):
            self.focus(target)
        self.press_key(ESCAPE)

    def _wait_closed(self, popup: Node) -> bool:
        return bool(self._waiter.until(lambda: self._gone(popup), CLOSE_WAIT))

    def _gone(self, popup: Node) -> bool:
        return popup.handle is None or self._app.attribute(popup.handle, "AXRole") is None

    def _window(self) -> object:
        return self._app.window_or_launch(
            lambda window: self._app.text(window, "AXSubrole") == STANDARD_WINDOW
        )

    def _settle(self) -> Node:
        latest: list[Node] = []

        def sample() -> int:
            node = self._app.snapshot(self._window())
            latest[:] = [node]
            return sum(1 for _ in node.walk())

        self._waiter.settle(sample, minimum=SETTLED_MIN_NODES, timeout=SETTLE_TIMEOUT)
        return latest[0]
