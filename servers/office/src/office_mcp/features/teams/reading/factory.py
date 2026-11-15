from __future__ import annotations

from office_mcp.features.teams.activity.service import ActivityReader
from office_mcp.features.teams.channels.service import ChannelReader
from office_mcp.features.teams.driver import TeamsDriver
from office_mcp.features.teams.messages.service import MessageService
from office_mcp.features.teams.reading.service import TeamsReaders, TeamsReadingService
from office_mcp.features.teams.reading.unread import UnreadKeeper
from office_mcp.features.teams.search.service import SearchReader
from office_mcp.shared.runtime import Runtime


def build_teams_readers(runtime: Runtime, teams: TeamsDriver) -> TeamsReaders:
    gateway = teams.gateway
    waiter = teams.waiter
    return TeamsReaders(
        chats=teams.chats,
        messages=MessageService(gateway, teams.opener, teams.zone, waiter),
        unread=UnreadKeeper(gateway, teams.chats, waiter),
        activity=ActivityReader(gateway, waiter),
        search=SearchReader(gateway, waiter, runtime.today),
        channels=ChannelReader(gateway, waiter, teams.zone),
    )


def build_teams_reading(runtime: Runtime, teams: TeamsDriver) -> TeamsReadingService:
    return TeamsReadingService(build_teams_readers(runtime, teams), teams.scope, runtime.today)
