from __future__ import annotations

import inspect
from datetime import date

import pytest
from fakes.clock import FakeClock
from fakes.drivers import fake_runtime, teams_rig
from fakes.guard import RecordingGuard
from fakes.lock import MemoryUiLock
from fakes.p4_calendar import CalendarDesk, CalendarRig, DeskEvent, calendar_rig
from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import teams_window
from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.features.calendar.ax import AxCalendar, refuse_unsafe
from office_mcp.features.calendar.factory import build_calendar
from office_mcp.features.calendar.service import CalendarService
from office_mcp.features.calendar.tools import CalendarTools
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.errors import ForbiddenActionError
from office_mcp.shared.events import Key
from office_mcp.shared.node import Node
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
)
from office_mcp.shared.runner import Lazy, Runner

MONDAY = date(2026, 11, 16)
OWNER = Owner("Alex Example")


def tools_for(rig: CalendarRig) -> CalendarTools:
    runner = Runner(MemoryUiLock(), FakeClock(), lock_timeout=20, stale_after=10)
    return CalendarTools(runner, Lazy(lambda: rig.service), OWNER)


def desk_rig() -> CalendarRig:
    return calendar_rig(CalendarDesk([DeskEvent("Project Atlas budget", MONDAY)]))


def test_specs_names_groups_effects_and_stability() -> None:
    specs = {spec.name: spec for spec in tools_for(desk_rig()).specs()}

    assert list(specs) == ["calendar_events", "calendar_event", "calendar_respond"]
    assert (specs["calendar_events"].group, specs["calendar_events"].effect) == (
        ToolGroup.READ,
        Effect.READS,
    )
    assert specs["calendar_event"].stability is Stability.STABLE
    respond = specs["calendar_respond"]
    assert (respond.group, respond.effect, respond.stability) == (
        ToolGroup.SEND,
        Effect.SENDS,
        Stability.EXPERIMENTAL,
    )
    assert respond.description.startswith(EXPERIMENTAL_PREFIX)
    assert all("Alex Example's" in spec.description for spec in specs.values())


def test_descriptions_are_dash_free() -> None:
    for spec in tools_for(desk_rig()).specs():
        assert "\u2013" not in spec.description
        assert "\u2014" not in spec.description


def test_tool_parameters_match_the_catalogue() -> None:
    tools = tools_for(desk_rig())

    assert list(inspect.signature(tools.calendar_events).parameters) == [
        "start",
        "days",
        "include_canceled",
        "limit",
    ]
    assert list(inspect.signature(tools.calendar_respond).parameters) == [
        "event",
        "response",
        "notify",
        "comment",
        "dry_run",
    ]


def test_tools_return_records_through_the_runner() -> None:
    rig = desk_rig()
    tools = tools_for(rig)

    events = tools.calendar_events(start="2026-11-16")
    details = tools.calendar_event(str(events[0]["id"]))
    dry = tools.calendar_respond(str(events[0]["id"]), "accept", dry_run=True)

    assert events[0]["subject"] == "Project Atlas budget"
    assert events[0]["start"] == "2026-11-16T10:00:00+00:00"
    assert details["my_response"] == "not_responded"
    assert dry["ok"] is True and dry["verified"] is False


def test_errors_become_tool_errors() -> None:
    tools = tools_for(desk_rig())

    with pytest.raises(ToolError, match="id from calendar_events"):
        tools.calendar_event("Project Atlas")
    with pytest.raises(ToolError, match="days"):
        tools.calendar_events(days=40)


def test_never_presses_join_or_chat() -> None:
    for label in ("Join", "Chat with participants", "Join with an ID"):
        with pytest.raises(ForbiddenActionError):
            refuse_unsafe(Node(role="AXButton", title=label))

    refuse_unsafe(
        Node(
            role="AXButton",
            title="Join us, 10:00 to 10:30, Monday, November 16, 2026, By Jane Doe, Busy",
        )
    )


def test_card_closes_with_escape_when_popups_leave_it_open() -> None:
    desk = CalendarDesk(
        [DeskEvent("Project Atlas budget", MONDAY)], app=TeamsApp.CALENDAR, peek_for=0
    )
    rig = calendar_rig(desk)
    gateway = AxCalendar(desk, rig.text, rig.app.app.waiter)

    assert gateway.close_popups() == 1
    assert desk.peek_for is None
    assert [key.key for key in desk.keys] == [Key.ESCAPE]
    assert desk.focused == ["view"]


def test_card_closed_by_popups_needs_no_escape() -> None:
    desk = CalendarDesk(
        [DeskEvent("Project Atlas budget", MONDAY)],
        app=TeamsApp.CALENDAR,
        peek_for=0,
        peek_ignores_popups=False,
    )
    rig = calendar_rig(desk)
    gateway = AxCalendar(desk, rig.text, rig.app.app.waiter)

    assert gateway.close_popups() == 1
    assert desk.keys == []


def test_factory_builds_a_service_from_the_teams_driver() -> None:
    teams = teams_rig(FakeTeamsGateway([teams_window("Chat", [], [])])).driver

    service = build_calendar(fake_runtime(), teams, RecordingGuard())

    assert isinstance(service, CalendarService)
