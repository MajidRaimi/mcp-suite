from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.writing.targets import ChatTargets
from office_mcp.shared.errors import CleanupError, MenuItemDisabledError, OfficeError, error_text
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.menus import find_item, refuse_forbidden
from office_mcp.shared.node import Node
from office_mcp.shared.recipients import RecipientGuard
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import ScopeFactory, ScopeOptions, UiScope
from office_mcp.shared.waiting import Waiter

WRITE_SCOPE = ScopeOptions(foreground=False, restore_view=True)
MIN_WAIT = 1.0


@dataclass(frozen=True)
class Scoped[T]:
    value: T
    focus: FocusOutcome
    notes: tuple[str, ...]


@dataclass(frozen=True)
class WriteContext:
    targets: ChatTargets
    gateway: TeamsGateway
    scopes: ScopeFactory
    waiter: Waiter
    recipients: RecipientGuard

    def run[T](self, work: Callable[[UiScope], T]) -> Scoped[T]:
        return run_scoped(self.scopes, WRITE_SCOPE, work)

    def menu_item(
        self, body: Node, titles: tuple[str, ...], missing: Callable[[], OfficeError]
    ) -> Node:
        menu = self.gateway.show_menu(body)
        found = (find_item(menu, (title,)) for title in titles)
        item = next((node for node in found if node is not None), None)
        if item is None:
            self.gateway.close_popups()
            raise missing()
        if item.enabled is False:
            self.gateway.close_popups()
            raise MenuItemDisabledError(f"the Teams menu item {item.label()!r} is disabled")
        refuse_forbidden((item.label(),))
        return item


def run_scoped[T](
    scopes: ScopeFactory, options: ScopeOptions, work: Callable[[UiScope], T]
) -> Scoped[T]:
    done: list[tuple[UiScope, T]] = []
    try:
        with scopes(options) as scope:
            done.append((scope, work(scope)))
    except CleanupError as error:
        if not done:
            raise
        scope, value = done[0]
        return Scoped(value=value, focus=scope.focus, notes=(error_text(error),))
    scope, value = done[0]
    return Scoped(value=value, focus=scope.focus, notes=())


def bounded(seconds: float, deadline: Deadline) -> float:
    return min(seconds, max(deadline.remaining(), MIN_WAIT))
