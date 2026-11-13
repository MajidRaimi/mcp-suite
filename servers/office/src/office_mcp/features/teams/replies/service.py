from __future__ import annotations

from datetime import date, timedelta

from office_mcp.features.teams.chats.models import Chat

SELF_CHAT_SUFFIX = "(You)"


def needs_reply(chats: list[Chat], today: date, days: int, include_meetings: bool) -> list[Chat]:
    since = today - timedelta(days=days)
    waiting = [
        chat
        for chat in chats
        if not chat.from_me
        and not chat.muted
        and chat.preview
        and not chat.name.endswith(SELF_CHAT_SUFFIX)
        and (include_meetings or chat.kind != "meeting")
        and chat.last_active is not None
        and date.fromisoformat(chat.last_active) >= since
    ]
    newest_first = sorted(waiting, key=lambda chat: chat.last_active or "", reverse=True)
    return sorted(newest_first, key=lambda chat: not chat.unread)
