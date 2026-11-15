from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from typing import Protocol


@dataclass(frozen=True)
class SentMail:
    subject: str
    body: str
    replied_to: str = ""
    people: tuple[str, ...] = ()


@dataclass(frozen=True)
class OwnMessage:
    chat: str
    text: str
    replied_to: str = ""
    people: tuple[str, ...] = ()


class SentMailSource(Protocol):
    def sent_mail(self, limit: int) -> Sequence[SentMail]: ...


class OwnChatSource(Protocol):
    def own_messages(self, chats: int, per_chat: int) -> Sequence[OwnMessage]: ...
