from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, tzinfo

from whatsapp_mcp.features.waha.client import WahaClient
from whatsapp_mcp.features.waha.models import ChatSummary, Message
from whatsapp_mcp.shared.errors import ChatNotFoundError

_OVERVIEW_LIMIT = 300
_UNANSWERED_SCAN = 25
_PREVIEW = 160


@dataclass(frozen=True)
class WaitingChat:
    chat_id: str
    name: str
    is_group: bool
    waiting: str
    last_message_at: str
    unanswered_messages: int
    last_message: str


@dataclass(frozen=True)
class ChatLine:
    chat_id: str
    name: str
    is_group: bool
    last_message_at: str
    last_from_me: bool
    last_message: str


@dataclass(frozen=True)
class MessageLine:
    at: str
    sender: str
    text: str


@dataclass(frozen=True)
class SearchHit:
    chat_id: str
    chat: str
    at: str
    sender: str
    text: str


class InboxService:
    def __init__(self, waha: WahaClient, zone: tzinfo) -> None:
        self._waha = waha
        self._zone = zone

    def needs_reply(
        self, now: float, limit: int, include_groups: bool, max_age_days: int
    ) -> list[WaitingChat]:
        oldest = now - max_age_days * 86_400
        candidates = [
            chat
            for chat in self._waha.chats(_OVERVIEW_LIMIT)
            if chat.last_message is not None
            and not chat.last_message.from_me
            and chat.last_message.timestamp >= oldest
            and (include_groups or not chat.is_group)
        ]
        candidates.sort(key=_last_timestamp, reverse=True)
        waiting = [self._waiting(chat, now) for chat in candidates[:limit]]
        return sorted(waiting, key=lambda item: item.last_message_at)

    def recent_chats(self, limit: int) -> list[ChatLine]:
        chats = sorted(self._waha.chats(limit), key=_last_timestamp, reverse=True)
        return [
            ChatLine(
                chat_id=chat.id,
                name=chat.display_name,
                is_group=chat.is_group,
                last_message_at=self._format_time(_last_timestamp(chat)),
                last_from_me=bool(chat.last_message and chat.last_message.from_me),
                last_message=_preview(chat.last_message),
            )
            for chat in chats
        ]

    def read_chat(self, chat: str, limit: int) -> tuple[ChatSummary, list[MessageLine]]:
        summary = self.resolve(chat)
        messages = sorted(self._waha.messages(summary.id, limit), key=lambda item: item.timestamp)
        return summary, [self._line(message, summary) for message in messages]

    def search(self, query: str, now: float, days: int, chat_limit: int) -> list[SearchHit]:
        needle = query.casefold()
        oldest = now - days * 86_400
        hits: list[SearchHit] = []
        chats = [
            chat
            for chat in sorted(self._waha.chats(_OVERVIEW_LIMIT), key=_last_timestamp, reverse=True)
            if _last_timestamp(chat) >= oldest
        ][:chat_limit]
        for chat in chats:
            if needle in chat.display_name.casefold():
                hits.append(
                    SearchHit(
                        chat.id,
                        chat.display_name,
                        self._format_time(_last_timestamp(chat)),
                        "",
                        "[chat name matches]",
                    )
                )
            for message in self._waha.messages(chat.id, 100):
                if message.timestamp >= oldest and needle in (message.body or "").casefold():
                    line = self._line(message, chat)
                    hits.append(
                        SearchHit(chat.id, chat.display_name, line.at, line.sender, line.text)
                    )
        return sorted(hits, key=lambda hit: hit.at, reverse=True)

    def resolve(self, chat: str) -> ChatSummary:
        chats = self._waha.chats(_OVERVIEW_LIMIT)
        exact = [item for item in chats if item.id == chat]
        if exact:
            return exact[0]
        wanted = chat.casefold().strip()
        named = [item for item in chats if item.display_name.casefold() == wanted]
        partial = named or [item for item in chats if wanted in item.display_name.casefold()]
        digits = "".join(character for character in chat if character.isdigit())
        if not partial and len(digits) >= 7:
            partial = [item for item in chats if item.id.startswith(digits)]
        if len(partial) == 1:
            return partial[0]
        if not partial:
            raise ChatNotFoundError(f"no chat matches {chat!r}")
        options = ", ".join(f"{item.display_name} ({item.id})" for item in partial[:8])
        raise ChatNotFoundError(f"{chat!r} matches several chats: {options}. Use the chat_id.")

    def _waiting(self, chat: ChatSummary, now: float) -> WaitingChat:
        timestamp = _last_timestamp(chat)
        unanswered = 1
        if not chat.is_group:
            recent = sorted(
                self._waha.messages(chat.id, _UNANSWERED_SCAN),
                key=lambda item: item.timestamp,
                reverse=True,
            )
            unanswered = 0
            for message in recent:
                if message.from_me:
                    break
                unanswered += 1
        return WaitingChat(
            chat_id=chat.id,
            name=chat.display_name,
            is_group=chat.is_group,
            waiting=_duration(now - timestamp),
            last_message_at=self._format_time(timestamp),
            unanswered_messages=max(unanswered, 1),
            last_message=_preview(chat.last_message),
        )

    def _line(self, message: Message, chat: ChatSummary) -> MessageLine:
        text = (message.body or "").strip() or ("[media]" if message.has_media else "[no text]")
        return MessageLine(self._format_time(message.timestamp), _sender(message, chat), text)

    def _format_time(self, timestamp: int) -> str:
        if not timestamp:
            return ""
        return datetime.fromtimestamp(timestamp, tz=self._zone).strftime("%Y-%m-%d %H:%M")


def _last_timestamp(chat: ChatSummary) -> int:
    return chat.last_message.timestamp if chat.last_message else 0


def _preview(message: Message | None) -> str:
    if message is None:
        return ""
    text = (message.body or "").strip() or ("[media]" if message.has_media else "[no text]")
    collapsed = " ".join(text.split())
    return collapsed if len(collapsed) <= _PREVIEW else collapsed[: _PREVIEW - 3] + "..."


def _sender(message: Message, chat: ChatSummary) -> str:
    if message.from_me:
        return "me"
    if chat.is_group:
        return (message.participant or message.author or "member").split("@")[0]
    return chat.display_name


def _duration(seconds: float) -> str:
    minutes = max(int(seconds // 60), 0)
    days, remainder = divmod(minutes, 1_440)
    hours, minutes = divmod(remainder, 60)
    if days:
        return f"{days}d {hours}h"
    if hours:
        return f"{hours}h {minutes}m"
    return f"{minutes}m"
