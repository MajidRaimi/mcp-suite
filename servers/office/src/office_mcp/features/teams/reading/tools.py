from __future__ import annotations

from collections.abc import Callable
from typing import Any, Literal

from office_mcp.features.teams.activity.models import ActivityFilter
from office_mcp.features.teams.reading import descriptions as text
from office_mcp.features.teams.reading.service import TeamsReadingService
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import Effect, ToolGroup, ToolSpec
from office_mcp.shared.runner import Deadline, Lazy, Runner

Record = dict[str, Any]
MAX_MESSAGES = 200
MAX_OLDER = 20
MAX_ACTIVITY = 100
MAX_SEARCH = 50
MAX_POSTS = 100


def clamp(value: int, low: int, high: int) -> int:
    return max(low, min(value, high))


class TeamsReadingTools:
    def __init__(self, runner: Runner, reading: Lazy[TeamsReadingService], owner: Owner) -> None:
        self._runner = runner
        self._reading = reading
        self._owner = owner

    def specs(self) -> list[ToolSpec]:
        return [
            self._spec(self.teams_chats, text.TEAMS_CHATS, Effect.READS),
            self._spec(self.teams_read, text.TEAMS_READ, Effect.MARKS_READ),
            self._spec(self.teams_needs_reply, text.teams_needs_reply(self._owner), Effect.READS),
            self._spec(self.teams_activity, text.TEAMS_ACTIVITY, Effect.MARKS_READ),
            self._spec(self.teams_search, text.TEAMS_SEARCH, Effect.READS),
            self._spec(self.teams_channels, text.TEAMS_CHANNELS, Effect.MARKS_READ),
            self._spec(self.teams_channel_read, text.TEAMS_CHANNEL_READ, Effect.MARKS_READ),
        ]

    def teams_chats(
        self, unread_only: bool = False, section: str | None = None, limit: int = 0
    ) -> list[Record]:
        return self._read(
            lambda _deadline: [
                chat.to_record() for chat in self._service().chats(unread_only, section, limit)
            ]
        )

    def teams_read(
        self, chat: str, limit: int = 30, older: int = 0, keep_unread: bool = False
    ) -> Record:
        count = clamp(limit, 1, MAX_MESSAGES) if limit > 0 else MAX_MESSAGES
        pages = clamp(older, 0, MAX_OLDER)
        return self._read(
            lambda deadline: (
                self._service().read(chat, count, pages, keep_unread, deadline).to_record()
            )
        )

    def teams_needs_reply(self, days: int = 7, include_meetings: bool = False) -> list[Record]:
        return self._read(
            lambda _deadline: [
                chat.to_record() for chat in self._service().needs_reply(days, include_meetings)
            ]
        )

    def teams_activity(
        self, filter: Literal["all", "unread", "mentions"] = "all", limit: int = 30
    ) -> list[Record]:
        wanted = ActivityFilter(filter)
        count = clamp(limit, 1, MAX_ACTIVITY)
        return self._read(
            lambda _deadline: [item.to_record() for item in self._service().activity(wanted, count)]
        )

    def teams_search(self, query: str, limit: int = 15) -> list[Record]:
        count = clamp(limit, 1, MAX_SEARCH)
        return self._read(
            lambda _deadline: [
                result.to_record() for result in self._service().search(query, count)
            ]
        )

    def teams_channels(self, team: str | None = None, unread_only: bool = False) -> list[Record]:
        return self._read(
            lambda _deadline: [
                channel.to_record() for channel in self._service().channels(team, unread_only)
            ]
        )

    def teams_channel_read(
        self, team: str, channel: str, limit: int = 20, replies: bool = True
    ) -> Record:
        count = clamp(limit, 1, MAX_POSTS)
        return self._read(
            lambda _deadline: (
                self._service().channel_read(team, channel, count, replies).to_record()
            )
        )

    def _service(self) -> TeamsReadingService:
        return self._reading.get()

    def _read[T](self, work: Callable[[Deadline], T]) -> T:
        return self._runner.read(work)

    @staticmethod
    def _spec(function: Callable[..., object], description: str, effect: Effect) -> ToolSpec:
        return ToolSpec(
            name=function.__name__,
            function=function,
            description=description,
            group=ToolGroup.READ,
            effect=effect,
        )
