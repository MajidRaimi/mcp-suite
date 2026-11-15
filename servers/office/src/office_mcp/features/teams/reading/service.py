from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import date
from typing import Any

from office_mcp.features.teams.activity.models import ActivityFilter, ActivityItem
from office_mcp.features.teams.activity.service import ActivityReader
from office_mcp.features.teams.channels.models import Channel, ChannelPosts
from office_mcp.features.teams.channels.service import ChannelReader
from office_mcp.features.teams.chats.models import Chat
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.messages.models import Conversation
from office_mcp.features.teams.messages.service import MessageService
from office_mcp.features.teams.reading.unread import UnreadKeeper
from office_mcp.features.teams.reading.view import ReadScopeFactory
from office_mcp.features.teams.replies.service import needs_reply
from office_mcp.features.teams.search.models import SearchResult
from office_mcp.features.teams.search.service import SearchReader, checked_query
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import ScopeOptions

QUIET = ScopeOptions(foreground=False, restore_view=False)
RESTORING = ScopeOptions(foreground=False, restore_view=True)


@dataclass(frozen=True)
class ChatReading:
    conversation: Conversation
    restored_unread: bool
    notes: tuple[str, ...] = ()

    def to_record(self) -> dict[str, Any]:
        return {
            **self.conversation.to_record(),
            "restored_unread": self.restored_unread,
            "notes": list(self.notes),
        }


@dataclass(frozen=True)
class TeamsReaders:
    chats: ChatService
    messages: MessageService
    unread: UnreadKeeper
    activity: ActivityReader
    search: SearchReader
    channels: ChannelReader


class TeamsReadingService:
    def __init__(
        self, readers: TeamsReaders, scope: ReadScopeFactory, today: Callable[[], date]
    ) -> None:
        self._readers = readers
        self._scope = scope
        self._today = today

    def chats(self, unread_only: bool, section: str | None, limit: int) -> list[Chat]:
        with self._scope(QUIET):
            return self._readers.chats.filtered(unread_only, section, limit)

    def read(
        self, chat: str, limit: int, older: int, keep_unread: bool, deadline: Deadline | None = None
    ) -> ChatReading:
        with self._scope(QUIET):
            read = self._readers.messages.open_and_read(chat, limit, older, deadline)
            if not keep_unread:
                return ChatReading(conversation=read.conversation, restored_unread=False)
            outcome = self._readers.unread.keep(read)
            return ChatReading(read.conversation, outcome.restored, outcome.notes)

    def needs_reply(self, days: int, include_meetings: bool) -> list[Chat]:
        with self._scope(QUIET):
            chats = self._readers.chats.filtered(False, None, 0)
        return needs_reply(chats, self._today(), days, include_meetings)

    def activity(self, wanted: ActivityFilter, limit: int) -> list[ActivityItem]:
        with self._scope(RESTORING) as scope:
            return self._readers.activity.read(scope, wanted, limit)

    def search(self, query: str, limit: int) -> list[SearchResult]:
        text = checked_query(query)
        with self._scope(QUIET) as scope:
            return self._readers.search.search(scope, text, limit)

    def channels(self, team: str | None, unread_only: bool) -> list[Channel]:
        with self._scope(RESTORING) as scope:
            rows = self._readers.channels.channels(scope, team, unread_only)
        return [row.channel for row in rows]

    def channel_read(self, team: str, channel: str, limit: int, replies: bool) -> ChannelPosts:
        with self._scope(RESTORING) as scope:
            return self._readers.channels.read(scope, team, channel, limit, replies)
