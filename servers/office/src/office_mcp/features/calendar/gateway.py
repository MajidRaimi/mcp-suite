from __future__ import annotations

from collections.abc import Callable
from contextlib import AbstractContextManager
from typing import Protocol

from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.node import Node
from office_mcp.shared.scope import ScopeOptions


class CalendarGateway(Protocol):
    def snapshot(self) -> Node: ...

    def press(self, node: Node) -> None: ...

    def close_popups(self) -> int: ...

    def write_text(self, node: Node, text: str) -> None: ...

    def read_text(self, node: Node, placeholder: str) -> str: ...

    def clear_text(self, node: Node, inserted: str, placeholder: str) -> bool: ...


class CalendarScope(Protocol):
    cleanup: CleanupStack

    def ensure_front(self) -> None: ...

    @property
    def focus(self) -> FocusOutcome: ...

    def show_app(self, app: TeamsApp) -> Node: ...


CalendarScopeFactory = Callable[[ScopeOptions], AbstractContextManager[CalendarScope]]
