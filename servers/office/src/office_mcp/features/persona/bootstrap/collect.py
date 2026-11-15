from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, replace

from office_mcp.features.persona.bootstrap.sources import (
    OwnChatSource,
    OwnMessage,
    SentMail,
    SentMailSource,
)
from office_mcp.shared.drafts import canonical_text
from office_mcp.shared.errors import PersonaError

MAX_EMAILS = 200
MAX_CHATS = 50
MAX_PER_CHAT = 100


@dataclass(frozen=True)
class Corpus:
    emails: tuple[SentMail, ...]
    messages: tuple[OwnMessage, ...]


def collect_corpus(
    mail: SentMailSource, chats: OwnChatSource, emails: int, chats_limit: int, per_chat: int
) -> Corpus:
    require_range("emails", emails, MAX_EMAILS)
    require_range("chats", chats_limit, MAX_CHATS)
    require_range("per_chat", per_chat, MAX_PER_CHAT)
    sent = mail.sent_mail(emails) if emails else ()
    own = chats.own_messages(chats_limit, per_chat) if chats_limit and per_chat else ()
    return Corpus(
        emails=tuple(dedupe_mail(sent))[:emails],
        messages=tuple(dedupe_messages(own))[: chats_limit * per_chat],
    )


def require_range(name: str, value: int, maximum: int) -> None:
    if not 0 <= value <= maximum:
        raise PersonaError(f"{name} must be between 0 and {maximum}, got {value}")


def dedupe_mail(items: Sequence[SentMail]) -> list[SentMail]:
    seen: set[str] = set()
    kept: list[SentMail] = []
    for item in items:
        body = canonical_text(item.body)
        if body and body not in seen:
            seen.add(body)
            kept.append(replace(item, subject=item.subject.strip(), body=body))
    return kept


def dedupe_messages(items: Sequence[OwnMessage]) -> list[OwnMessage]:
    seen: set[tuple[str, str]] = set()
    kept: list[OwnMessage] = []
    for item in items:
        text = canonical_text(item.text)
        if text and (item.chat, text) not in seen:
            seen.add((item.chat, text))
            kept.append(replace(item, text=text))
    return kept
