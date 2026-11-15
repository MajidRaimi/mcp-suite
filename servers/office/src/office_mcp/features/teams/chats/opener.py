from __future__ import annotations

from dataclasses import dataclass

from office_mcp.features.teams.chats.models import ChatRow
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.header import ThreadKind, open_thread_id, thread_kind
from office_mcp.features.teams.messages.parse import message_pane, open_chat_name
from office_mcp.shared.errors import (
    AmbiguousTargetError,
    ForbiddenActionError,
    TargetError,
    UiTimeoutError,
)
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

WRITE_TARGET_HINT = "use the thread id or the exact chat name from teams_chats"
CHANNEL_REFUSAL = "office-mcp never writes in Teams channels or other teams' threads"
MAX_CANDIDATES = 8


@dataclass(frozen=True)
class OpenedChat:
    row: ChatRow
    root: Node
    thread_id: str


def refuse_channel(thread_id: str) -> None:
    if thread_kind(thread_id) is ThreadKind.CHANNEL:
        raise ForbiddenActionError(CHANNEL_REFUSAL)


def resolve_for_write(query: str, rows: list[ChatRow]) -> ChatRow:
    wanted = query.strip()
    refuse_channel(wanted)
    by_id = [row for row in rows if row.chat.thread_id == wanted]
    if by_id:
        refuse_channel(by_id[0].chat.thread_id)
        return by_id[0]
    name = normalize(wanted)
    by_name = [row for row in rows if name and normalize(row.chat.name) == name]
    if len(by_name) > 1:
        names = [row.chat.name for row in by_name[:MAX_CANDIDATES]]
        raise AmbiguousTargetError(f"{query!r} names several chats; use the thread id", names)
    if not by_name:
        raise TargetError(WRITE_TARGET_HINT)
    refuse_channel(by_name[0].chat.thread_id)
    return by_name[0]


def shows_thread(root: Node, row: ChatRow) -> bool:
    opened = open_thread_id(root)
    if opened is not None:
        return opened == row.chat.thread_id
    return open_chat_name(root) == row.chat.name


class ChatOpener:
    def __init__(self, gateway: TeamsGateway, chats: ChatService, waiter: Waiter) -> None:
        self._gateway = gateway
        self._chats = chats
        self._waiter = waiter

    def open(self, query: str, for_write: bool) -> OpenedChat:
        rows = self._chats.rows()
        row = resolve_for_write(query, rows) if for_write else self._chats.resolve(query, rows)
        return self.open_row(row)

    def open_row(self, row: ChatRow) -> OpenedChat:
        root = self._gateway.snapshot()
        if shows_thread(root, row) and message_pane(root) is not None:
            return OpenedChat(row=row, root=root, thread_id=row.chat.thread_id)
        self._gateway.press(row.node)
        self._waiter.require(
            lambda: shows_thread(self._gateway.snapshot(), row),
            lambda: UiTimeoutError(f"Teams did not open the chat {row.chat.name!r} in time"),
        )
        opened = self._waiter.until(lambda: self._with_pane(row)) or self._gateway.snapshot()
        return OpenedChat(row=row, root=opened, thread_id=row.chat.thread_id)

    def find_self(self, thread_id: str) -> ChatRow | None:
        return next((row for row in self._chats.rows() if row.chat.thread_id == thread_id), None)

    def _with_pane(self, row: ChatRow) -> Node | None:
        root = self._gateway.snapshot()
        if shows_thread(root, row) and message_pane(root) is not None:
            return root
        return None
