from __future__ import annotations

import time
from collections.abc import Callable
from datetime import tzinfo

from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.messages.models import Conversation
from office_mcp.features.teams.messages.parse import (
    first_content,
    message_pane,
    open_chat_name,
    parse_messages,
)
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.node import Node

POLL_SECONDS = 0.5


class MessageService:
    def __init__(
        self,
        gateway: TeamsGateway,
        chats: ChatService,
        zone: tzinfo,
        timeout: float = 12.0,
        sleep: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self._gateway = gateway
        self._chats = chats
        self._zone = zone
        self._timeout = timeout
        self._sleep = sleep
        self._clock = clock

    def read(self, query: str, limit: int, older: int) -> Conversation:
        rows = self._chats.rows()
        row = self._chats.resolve(query, rows)
        root = self._open(row.chat.name, row.node)
        for _ in range(max(older, 0)):
            loaded = self._load_older(root)
            if loaded is None:
                break
            root = loaded
        messages = parse_messages(root, self._zone)
        shown = messages[-limit:] if limit > 0 else messages
        return Conversation(chat=row.chat, messages=shown, loaded=len(messages))

    def _open(self, name: str, row: Node) -> Node:
        root = self._gateway.snapshot()
        if open_chat_name(root) == name and message_pane(root) is not None:
            return root
        self._gateway.press(row)
        opened = self._wait(lambda candidate: open_chat_name(candidate) == name)
        if opened is None:
            raise OfficeError(f"Teams did not open the chat {name!r} in time")
        return self._wait(lambda candidate: message_pane(candidate) is not None) or opened

    def _load_older(self, root: Node) -> Node | None:
        first = first_content(root)
        if first is None:
            return None
        self._gateway.scroll_into_view(first)
        return self._wait(lambda candidate: changed_first(candidate, first.dom_id))

    def _wait(self, ready: Callable[[Node], bool]) -> Node | None:
        deadline = self._clock() + self._timeout
        while True:
            root = self._gateway.snapshot()
            if ready(root):
                return root
            if self._clock() >= deadline:
                return None
            self._sleep(POLL_SECONDS)


def changed_first(root: Node, previous_id: str) -> bool:
    first = first_content(root)
    return first is not None and first.dom_id != previous_id
