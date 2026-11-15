from __future__ import annotations

from collections.abc import Iterable

from office_mcp.features.teams.messages.models import Message
from office_mcp.shared.text import normalize

DUPLICATE_LOOKBACK = 5
DELETED_MARKS = ("this message has been deleted", "message deleted")


def collapse(text: str) -> str:
    return " ".join(text.split())


def numeric_id(message: Message) -> int | None:
    return int(message.id) if message.id.isdigit() else None


def own_messages(messages: Iterable[Message]) -> list[Message]:
    return [message for message in messages if message.mine]


def newest_id(messages: Iterable[Message]) -> int:
    ids = [numeric_id(message) for message in messages]
    return max((value for value in ids if value is not None), default=0)


def newer_own(messages: Iterable[Message], before: int) -> list[Message]:
    return [
        message
        for message in own_messages(messages)
        if (value := numeric_id(message)) is not None and value > before
    ]


def sent_early(messages: Iterable[Message], before: int) -> bool:
    return bool(newer_own(messages, before))


def newest_from_others(messages: Iterable[Message]) -> Message | None:
    others = [message for message in messages if not message.mine]
    return others[-1] if others else None


def shows_text(message: Message, text: str, quoted: bool) -> bool:
    shown = collapse(message.text)
    wanted = collapse(text)
    if shown == wanted:
        return True
    return quoted and bool(wanted) and shown.endswith(" " + wanted)


def find_sent(
    messages: Iterable[Message], before: int, text: str, quoted: bool = False
) -> Message | None:
    return next(
        (message for message in newer_own(messages, before) if shows_text(message, text, quoted)),
        None,
    )


def recent_duplicate(
    messages: Iterable[Message], text: str, lookback: int = DUPLICATE_LOOKBACK
) -> Message | None:
    wanted = normalize(text)
    recent = own_messages(messages)[-lookback:]
    return next((message for message in recent if normalize(message.text) == wanted), None)


def looks_deleted(text: str) -> bool:
    folded = normalize(text)
    return any(mark in folded for mark in DELETED_MARKS)
