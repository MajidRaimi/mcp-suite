from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from enum import StrEnum
from typing import Protocol

from office_mcp.shared.node import Node


class RecipientField(StrEnum):
    TO = "To"
    CC = "Cc"
    BCC = "Bcc"


class ReplyKind(StrEnum):
    REPLY = "Reply"
    REPLY_ALL = "Reply all"
    FORWARD = "Forward"


@dataclass(frozen=True, eq=False)
class ComposeSurface:
    handle: object
    window: object
    inline: bool


@dataclass(frozen=True)
class RecipientToken:
    label: str
    address: str | None


@dataclass(frozen=True)
class RecipientRead:
    tokens: tuple[RecipientToken, ...]
    pending: str


@dataclass(frozen=True)
class AlertSheet:
    text: str
    buttons: tuple[str, ...]
    node: Node


class ComposeGateway(Protocol):
    def surfaces(self) -> list[ComposeSurface]: ...

    def stray_surfaces(self, before: Sequence[ComposeSurface]) -> list[ComposeSurface]: ...

    def open_new(self) -> ComposeSurface: ...

    def open_reply(self, row: Node, kind: ReplyKind) -> ComposeSurface: ...

    def dismiss_popovers(self) -> int: ...

    def body(self, surface: ComposeSurface) -> str: ...

    def insert_body(self, surface: ComposeSurface, text: str) -> None: ...

    def subject(self, surface: ComposeSurface) -> str: ...

    def set_subject(self, surface: ComposeSurface, text: str) -> None: ...

    def add_recipient(
        self, surface: ComposeSurface, field: RecipientField, address: str
    ) -> None: ...

    def recipients(self, surface: ComposeSurface, field: RecipientField) -> RecipientRead: ...

    def set_sensitivity(self, surface: ComposeSurface, path: tuple[str, ...]) -> None: ...

    def shown_texts(self, surface: ComposeSurface) -> list[str]: ...

    def send_button(self, surface: ComposeSurface) -> Node: ...

    def press_send(self, button: Node) -> None: ...

    def is_open(self, surface: ComposeSurface) -> bool: ...

    def alerts(self, window: object) -> list[AlertSheet]: ...

    def new_alert(self, window: object, before: Sequence[AlertSheet]) -> AlertSheet | None: ...

    def press_alert(self, alert: AlertSheet, title: str) -> None: ...

    def discard(self, surface: ComposeSurface) -> None: ...

    def undo_send_enabled(self) -> bool: ...

    def undo_send(self) -> None: ...
