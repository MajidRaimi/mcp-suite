from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any

from office_mcp.features.teams.chats.models import Chat


@dataclass(frozen=True)
class Quote:
    author: str
    sent: str
    text: str

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class Message:
    id: str
    author: str
    mine: bool
    sent_at: str
    text: str
    quote: Quote | None = None
    edited: bool = False

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class Conversation:
    chat: Chat
    messages: list[Message]
    loaded: int

    def to_record(self) -> dict[str, Any]:
        return {
            "chat": self.chat.to_record(),
            "loaded": self.loaded,
            "returned": len(self.messages),
            "messages": [message.to_record() for message in self.messages],
        }
