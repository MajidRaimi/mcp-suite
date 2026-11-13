from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any

from office_mcp.shared.node import Node


@dataclass(frozen=True)
class Chat:
    thread_id: str
    name: str
    kind: str
    section: str
    unread: bool
    muted: bool
    presence: str
    preview: str
    last_sender: str
    from_me: bool
    time_label: str
    last_active: str | None

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ChatRow:
    chat: Chat
    node: Node
