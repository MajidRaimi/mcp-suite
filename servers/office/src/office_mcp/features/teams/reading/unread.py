from __future__ import annotations

from dataclasses import dataclass

from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.messages.models import Message
from office_mcp.features.teams.messages.parse import message_body
from office_mcp.features.teams.messages.service import ChatRead
from office_mcp.shared.errors import (
    ActionNotVerifiedError,
    ElementNotFoundError,
    OfficeError,
    error_text,
)
from office_mcp.shared.menus import find_item
from office_mcp.shared.waiting import Waiter

MARK_UNREAD = ("Mark as unread",)
VERIFY_WAIT = 5.0
NOT_RESTORED = "Teams marked the chat read and it could not be marked unread again: "
NO_OTHER_MESSAGE = "the loaded messages hold none from others to mark unread"


@dataclass(frozen=True)
class UnreadOutcome:
    restored: bool
    notes: tuple[str, ...] = ()


def newest_from_others(messages: list[Message]) -> Message | None:
    return next((message for message in reversed(messages) if not message.mine), None)


class UnreadKeeper:
    def __init__(self, gateway: TeamsGateway, chats: ChatService, waiter: Waiter) -> None:
        self._gateway = gateway
        self._chats = chats
        self._waiter = waiter

    def keep(self, read: ChatRead) -> UnreadOutcome:
        if not read.row.chat.unread:
            return UnreadOutcome(restored=False)
        try:
            self._mark_unread(read)
        except OfficeError as error:
            return UnreadOutcome(restored=False, notes=(NOT_RESTORED + error_text(error),))
        return UnreadOutcome(restored=True)

    def _mark_unread(self, read: ChatRead) -> None:
        newest = newest_from_others(read.messages)
        if newest is None:
            raise ElementNotFoundError(NO_OTHER_MESSAGE)
        body = message_body(self._gateway.snapshot(), newest.id)
        if body is None:
            raise ElementNotFoundError(f"the message {newest.id} is no longer on screen")
        menu = self._gateway.show_menu(body)
        item = find_item(menu, MARK_UNREAD)
        if item is None:
            self._gateway.close_popups()
            raise ElementNotFoundError("the message menu has no Mark as unread item")
        self._gateway.press(item)
        self._gateway.close_popups()
        self._waiter.require(
            lambda: self._shows_unread(read.row.chat.thread_id),
            lambda: ActionNotVerifiedError("the chat list does not show the chat as unread"),
            VERIFY_WAIT,
        )

    def _shows_unread(self, thread_id: str) -> bool:
        return any(
            row.chat.thread_id == thread_id and row.chat.unread for row in self._chats.rows()
        )
