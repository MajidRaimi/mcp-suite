from __future__ import annotations

from contextlib import suppress
from typing import Protocol

from office_mcp.features.calendar.parse import parse_event_title
from office_mcp.features.calendar.peek import find_peek
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.shared.errors import ForbiddenActionError, MenuStuckError, OfficeError
from office_mcp.shared.events import Key, Shortcut
from office_mcp.shared.menus import is_forbidden
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

ESCAPE = Shortcut(Key.ESCAPE)
CLOSE_WAIT = 2.0
FOCUS_ROLES = frozenset({"AXButton", "AXCheckBox", "AXTextArea", "AXLink"})
NEVER_PRESSED = frozenset(
    {
        "join",
        "chat with participants",
        "join with an id",
        "start an instant teams meeting.",
        "start an instant teams meeting",
        "create a new event.",
        "create a new event",
    }
)


class TextPort(Protocol):
    def read(self, element: object, placeholder: str = "") -> str: ...

    def write(
        self, element: object, text: str, at_start: bool = False, placeholder: str = ""
    ) -> None: ...

    def clear_exact(self, element: object, inserted: str, placeholder: str = "") -> bool: ...


def refuse_unsafe(node: Node) -> None:
    label = node.label()
    if parse_event_title(label) is not None:
        return
    if normalize(label) in NEVER_PRESSED or is_forbidden((label,)):
        raise ForbiddenActionError(f"office-mcp never presses {label[:60]!r} in the Teams calendar")


class AxCalendar:
    def __init__(self, teams: TeamsGateway, text: TextPort, waiter: Waiter) -> None:
        self._teams = teams
        self._text = text
        self._waiter = waiter

    def snapshot(self) -> Node:
        return self._teams.snapshot()

    def press(self, node: Node) -> None:
        refuse_unsafe(node)
        self._teams.press(node)

    def close_popups(self) -> int:
        closed = self._teams.close_popups()
        peek = find_peek(self._teams.snapshot())
        if peek is None:
            return closed
        self._escape(peek)
        if self._waiter.until(self._peek_gone, CLOSE_WAIT) is None:
            raise MenuStuckError("the Teams event card did not close")
        return closed + 1

    def write_text(self, node: Node, text: str) -> None:
        self._text.write(self._teams.element(node), text)

    def read_text(self, node: Node, placeholder: str) -> str:
        return self._text.read(self._teams.element(node), placeholder)

    def clear_text(self, node: Node, inserted: str, placeholder: str) -> bool:
        return self._text.clear_exact(self._teams.element(node), inserted, placeholder)

    def _escape(self, peek: Node) -> None:
        target = peek.find(lambda node: node is not peek and node.role in FOCUS_ROLES) or peek
        with suppress(OfficeError):
            self._teams.focus(target)
        self._teams.press_key(ESCAPE)

    def _peek_gone(self) -> bool:
        return find_peek(self._teams.snapshot()) is None
