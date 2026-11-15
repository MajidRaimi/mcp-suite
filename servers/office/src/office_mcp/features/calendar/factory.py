from __future__ import annotations

from office_mcp.features.calendar.ax import AxCalendar
from office_mcp.features.calendar.service import CalendarDeps, CalendarService
from office_mcp.features.teams.driver import TeamsDriver
from office_mcp.shared.drafts import TextGuard
from office_mcp.shared.runtime import Runtime


def build_calendar(runtime: Runtime, teams: TeamsDriver, guard: TextGuard) -> CalendarService:
    deps = CalendarDeps(
        gateway=AxCalendar(teams.gateway, teams.text, teams.waiter),
        scope=teams.scope,
        waiter=teams.waiter,
        today=runtime.today,
        zone=teams.zone,
        recipients=runtime.recipients,
        guard=guard,
    )
    return CalendarService(deps)
