from __future__ import annotations

from datetime import date

import pytest
from fakes.clock import FakeClock
from fakes.p4_calendar import WEEK, WORK_WEEK, CalendarDesk, DeskEvent, calendar_rig, week_start

from office_mcp.features.calendar.models import CalendarEvent
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.errors import (
    AmbiguousTargetError,
    TargetError,
    TargetNotFoundError,
    UiTimeoutError,
)
from office_mcp.shared.runner import Deadline

SUNDAY = date(2026, 11, 15)
MONDAY = date(2026, 11, 16)
SATURDAY = date(2026, 11, 21)
NEXT_TUESDAY = date(2026, 11, 24)


def deadline(seconds: float = 120.0) -> Deadline:
    return Deadline.after(FakeClock(), seconds)


def week_events() -> list[DeskEvent]:
    return [
        DeskEvent("Project Atlas budget", MONDAY),
        DeskEvent("Weekend hike", SATURDAY, hours="all day event", location="", show_as="Free"),
        DeskEvent("Canceled: Vendor call", MONDAY, hours="13:00 to 14:00", show_as="Free"),
        DeskEvent("Quarterly review", NEXT_TUESDAY, hours="09:00 to 10:00"),
    ]


def subjects(events: list[CalendarEvent]) -> list[str]:
    return [event.subject for event in events]


def test_lists_one_week_in_week_view_and_restores_view_and_rail() -> None:
    rig = calendar_rig(CalendarDesk(week_events()))

    events = rig.service.events("2026-11-15", 7, False, 100, deadline())

    assert subjects(events) == ["Project Atlas budget", "Weekend hike"]
    assert rig.desk.view == WORK_WEEK
    assert rig.desk.app is TeamsApp.CHAT
    assert "view:Week" in rig.desk.pressed
    assert rig.desk.pressed[-2:] == ["view:Work week", "rail:Chat"]


def test_range_coverage_walks_into_the_next_week_and_back() -> None:
    rig = calendar_rig(CalendarDesk(week_events()))

    events = rig.service.events("2026-11-20", 6, False, 100, deadline())

    assert subjects(events) == ["Weekend hike", "Quarterly review"]
    assert rig.desk.pressed.count("next") == 1
    assert "today" in rig.desk.pressed
    assert rig.desk.week_first == week_start(SUNDAY)


def test_a_week_in_the_past_is_reached_with_previous_and_restored_by_walking() -> None:
    shown = week_start(date(2026, 11, 29))
    desk = CalendarDesk([DeskEvent("Old sync", date(2026, 11, 3))], week_first=shown)
    rig = calendar_rig(desk)

    events = rig.service.events("2026-11-03", 1, False, 100, deadline())

    assert subjects(events) == ["Old sync"]
    assert desk.pressed.count("prev") == 4
    assert desk.pressed.count("next") == 4
    assert "today" not in desk.pressed
    assert desk.week_first == shown


def test_canceled_events_only_when_asked_and_limit_applies() -> None:
    rig = calendar_rig(CalendarDesk(week_events()))

    events = rig.service.events("2026-11-16", 1, True, 1, deadline())

    assert subjects(events) == ["Project Atlas budget"]
    every = rig.service.events("2026-11-16", 1, True, 10, deadline())
    assert [event.canceled for event in every] == [False, True]


def test_week_view_already_selected_is_left_alone() -> None:
    desk = CalendarDesk(week_events(), view=WEEK)
    rig = calendar_rig(desk)

    rig.service.events(None, 1, False, 100, deadline())

    assert "picker" not in desk.pressed
    assert desk.view == WEEK


@pytest.mark.parametrize(
    ("start", "days", "limit", "message"),
    [
        ("not a date", 1, 10, "ISO date"),
        (None, 0, 10, "days"),
        (None, 32, 10, "days"),
        (None, 1, 0, "limit"),
        (None, 1, 301, "limit"),
    ],
)
def test_invalid_arguments_touch_nothing(
    start: str | None, days: int, limit: int, message: str
) -> None:
    rig = calendar_rig(CalendarDesk(week_events()))

    with pytest.raises(TargetError, match=message):
        rig.service.events(start, days, False, limit, deadline())

    assert rig.desk.snapshots == 0


def test_expired_deadline_stops_week_navigation() -> None:
    rig = calendar_rig(CalendarDesk(week_events()))

    with pytest.raises(UiTimeoutError, match="fewer days"):
        rig.service.events("2026-11-20", 6, False, 100, deadline(0.0))

    assert rig.desk.app is TeamsApp.CHAT


def test_event_details_open_and_close_the_card() -> None:
    rig = calendar_rig(CalendarDesk(week_events()))
    listed = rig.service.events("2026-11-16", 1, False, 100, deadline())

    details = rig.service.event(listed[0].id, deadline()).to_record()

    assert details["subject"] == "Project Atlas budget"
    assert details["organizer"] == "Jane Doe"
    assert details["my_response"] == "not_responded"
    assert details["has_join_button"] is True
    assert rig.desk.peek_for is None
    assert "join" not in rig.desk.pressed
    assert "chat" not in rig.desk.pressed


def test_event_in_another_week_is_found_through_the_remembered_day() -> None:
    rig = calendar_rig(CalendarDesk(week_events()))
    listed = rig.service.events("2026-11-24", 1, False, 100, deadline())

    details = rig.service.event(listed[0].id, deadline())

    assert details.event.subject == "Quarterly review"
    assert rig.desk.week_first == week_start(SUNDAY)


def test_id_collision_raises_ambiguity() -> None:
    twin = DeskEvent("Project Atlas budget", MONDAY)
    rig = calendar_rig(CalendarDesk([DeskEvent("Project Atlas budget", MONDAY), twin]))
    listed = rig.service.events("2026-11-16", 1, False, 100, deadline())

    assert listed[0].id == listed[1].id
    with pytest.raises(AmbiguousTargetError) as caught:
        rig.service.event(listed[0].id, deadline())

    assert caught.value.candidates == ("Project Atlas budget", "Project Atlas budget")
    assert rig.desk.peek_for is None


def test_unknown_and_malformed_ids() -> None:
    rig = calendar_rig(CalendarDesk(week_events()))

    with pytest.raises(TargetNotFoundError, match="calendar_events"):
        rig.service.event("0123456789ab", deadline())
    with pytest.raises(TargetError, match="id from calendar_events"):
        rig.service.event("Project Atlas", deadline())

    assert rig.desk.snapshots == 0


def test_vanished_event_is_reported() -> None:
    desk = CalendarDesk(week_events())
    rig = calendar_rig(desk)
    listed = rig.service.events("2026-11-16", 1, False, 100, deadline())
    desk.events.pop(0)

    with pytest.raises(TargetNotFoundError, match="no longer"):
        rig.service.event(listed[0].id, deadline())


def test_the_week_comes_from_the_range_button_not_the_week_navigation_buttons() -> None:
    rig = calendar_rig(CalendarDesk(week_events()))

    events = rig.service.events("2026-11-16", 1, False, 100, deadline())

    assert subjects(events) == ["Project Atlas budget"]
    assert "prev" not in rig.desk.pressed
    assert "next" not in rig.desk.pressed


def test_a_view_picker_that_renders_late_is_waited_for() -> None:
    rig = calendar_rig(CalendarDesk(week_events(), picker_delay=3))

    events = rig.service.events("2026-11-15", 7, False, 100, deadline())

    assert subjects(events) == ["Project Atlas budget", "Weekend hike"]
