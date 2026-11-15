from __future__ import annotations

from dataclasses import dataclass
from datetime import tzinfo

from office_mcp.features.teams.chats.models import ChatRow
from office_mcp.features.teams.chats.opener import ChatOpener
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.messages.models import Conversation, Message
from office_mcp.features.teams.messages.parse import first_content, parse_messages
from office_mcp.shared.node import Node
from office_mcp.shared.runner import Deadline
from office_mcp.shared.waiting import Waiter


@dataclass(frozen=True)
class ChatRead:
    row: ChatRow
    root: Node
    messages: list[Message]
    conversation: Conversation


class MessageService:
    def __init__(
        self, gateway: TeamsGateway, opener: ChatOpener, zone: tzinfo, waiter: Waiter
    ) -> None:
        self._gateway = gateway
        self._opener = opener
        self._zone = zone
        self._waiter = waiter

    def read(self, query: str, limit: int, older: int) -> Conversation:
        return self.open_and_read(query, limit, older).conversation

    def open_and_read(
        self, query: str, limit: int, older: int, deadline: Deadline | None = None
    ) -> ChatRead:
        opened = self._opener.open(query, for_write=False)
        root = opened.root
        for _ in range(max(older, 0)):
            if deadline is not None and deadline.expired():
                break
            loaded = self._load_older(root)
            if loaded is None:
                break
            root = loaded
        messages = parse_messages(root, self._zone)
        shown = messages[-limit:] if limit > 0 else messages
        conversation = Conversation(chat=opened.row.chat, messages=shown, loaded=len(messages))
        return ChatRead(row=opened.row, root=root, messages=messages, conversation=conversation)

    def _load_older(self, root: Node) -> Node | None:
        first = first_content(root)
        if first is None:
            return None
        self._gateway.scroll_into_view(first)
        return self._waiter.until(lambda: self._changed_first(first.dom_id))

    def _changed_first(self, previous_id: str) -> Node | None:
        root = self._gateway.snapshot()
        return root if changed_first(root, previous_id) else None


def changed_first(root: Node, previous_id: str) -> bool:
    first = first_content(root)
    return first is not None and first.dom_id != previous_id
