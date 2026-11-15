from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from office_mcp.shared.events import Shortcut
from office_mcp.shared.node import Node


@dataclass(frozen=True)
class MoveSheet:
    handle: object


@dataclass(frozen=True)
class MoveTarget:
    folder: str
    account: str
    description: str


class ActionsGateway(Protocol):
    def press_menu(self, path: tuple[str, ...], shortcut: Shortcut | None = None) -> None: ...

    def menu_titles(self, path: tuple[str, ...]) -> tuple[str, ...]: ...

    def open_row_menu(self) -> Node: ...

    def choose(self, menu: Node, path: tuple[str, ...]) -> None: ...

    def checked(self, menu: Node, path: tuple[str, ...]) -> bool | None: ...

    def close_menu(self, menu: Node | None = None) -> None: ...

    def open_move_sheet(self) -> MoveSheet: ...

    def filter_move_sheet(self, sheet: MoveSheet, text: str) -> tuple[MoveTarget, ...]: ...

    def pick_move_target(self, sheet: MoveSheet, target: MoveTarget) -> None: ...

    def confirm_move(self, sheet: MoveSheet) -> None: ...

    def cancel_move(self, sheet: MoveSheet) -> None: ...

    def sheet_open(self, sheet: MoveSheet) -> bool: ...
