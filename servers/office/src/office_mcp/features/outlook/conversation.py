from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from office_mcp.features.outlook.models import Person


@dataclass(frozen=True)
class LatestMessage:
    key: str
    subject: str
    sender: Person | None
    to: tuple[Person, ...]
    cc: tuple[Person, ...]
    hidden_recipients: int
    body_top: str
    sent: str
    folder: str


class ConversationReader(Protocol):
    def latest_message(self, key: str, folder: str | None) -> LatestMessage: ...
