from __future__ import annotations

from datetime import UTC, date, datetime, timedelta, timezone

import pytest

from office_mcp.features.calendar.models import DateRange
from office_mcp.features.calendar.parse import (
    build_event,
    event_id,
    parse_event_title,
    parse_range_label,
    parse_view_label,
)

UTC_PLUS_THREE = timezone(timedelta(hours=3))
MEETING = (
    "Project Atlas budget, 10:00 to 10:30, Monday, November 16, 2026, "
    "Microsoft Teams Meeting, By Jane Doe, Busy"
)


@pytest.mark.parametrize(
    ("title", "subject", "location", "organizer", "show_as", "recurring", "canceled"),
    [
        (
            MEETING,
            "Project Atlas budget",
            "Microsoft Teams Meeting",
            "Jane Doe",
            "busy",
            False,
            False,
        ),
        (
            "Weekly sync, 09:00 to 09:30, Tuesday, November 17, 2026, By Omar Haddad, Tentative",
            "Weekly sync",
            "",
            "Omar Haddad",
            "tentative",
            False,
            False,
        ),
        (
            "Standup, 08:30 to 08:45, Wednesday, November 18, 2026, Room 4, By Alex Example, "
            "Free, Recurring event",
            "Standup",
            "Room 4",
            "Alex Example",
            "free",
            True,
            False,
        ),
        (
            "Canceled: Vendor call, 13:00 to 14:00, Thursday, November 19, 2026, "
            "Microsoft Teams Meeting, By Jane Doe, Free",
            "Vendor call",
            "Microsoft Teams Meeting",
            "Jane Doe",
            "free",
            False,
            True,
        ),
        (
            "Review, plan, and ship, 15:00 to 16:00, Friday, November 20, 2026, "
            "Building 2, Floor 3, By Doe, Jane, Away",
            "Review, plan, and ship",
            "Building 2, Floor 3",
            "Doe, Jane",
            "out_of_office",
            False,
            False,
        ),
    ],
)
def test_title_parser_table(
    title: str,
    subject: str,
    location: str,
    organizer: str,
    show_as: str,
    recurring: bool,
    canceled: bool,
) -> None:
    parsed = parse_event_title(title)

    assert parsed is not None
    assert (parsed.subject, parsed.location, parsed.organizer) == (subject, location, organizer)
    assert (parsed.show_as, parsed.recurring, parsed.canceled) == (show_as, recurring, canceled)


def test_all_day_event_spans_the_whole_day() -> None:
    title = "Offsite, all day event, Saturday, November 21, 2026, By Jane Doe, Free"

    event = build_event(title, UTC)

    assert event is not None
    assert event.all_day is True
    assert event.to_record()["start"] == "2026-11-21"
    assert event.to_record()["end"] == "2026-11-22"


def test_times_carry_the_zone_offset_and_overnight_end_moves_to_next_day() -> None:
    title = "Release window, 23:00 to 01:00, Sunday, November 15, 2026, By Jane Doe, Busy"

    event = build_event(title, UTC_PLUS_THREE)

    assert event is not None
    assert event.start == datetime(2026, 11, 15, 23, 0, tzinfo=UTC_PLUS_THREE)
    assert event.end_text() == "2026-11-16T01:00:00+03:00"


def test_twelve_hour_times_with_narrow_spaces() -> None:
    title = "Lunch, 12:30\u202fPM to 1:15\u202fPM, Monday, November 16, 2026, By Jane Doe, Busy"

    event = build_event(title, UTC)

    assert event is not None
    assert (event.start.hour, event.start.minute, event.end.hour) == (12, 30, 13)


def test_online_meeting_follows_the_location() -> None:
    event = build_event(MEETING, UTC)

    assert event is not None
    assert event.online_meeting is True
    assert event.day == date(2026, 11, 16)


@pytest.mark.parametrize(
    "label",
    [
        "Go to today",
        "15, November, 2026",
        "Change the view of your calendar, button. Work week selected",
        "Project Atlas budget, 10:00, Monday",
    ],
)
def test_non_event_labels_do_not_parse(label: str) -> None:
    assert parse_event_title(label) is None


def test_id_is_twelve_hex_and_depends_on_subject_start_and_organizer() -> None:
    first = build_event(MEETING, UTC)
    moved = build_event(MEETING.replace("10:00 to", "11:00 to"), UTC)
    canceled = build_event("Canceled: " + MEETING, UTC)

    assert first is not None and moved is not None and canceled is not None
    assert first.id == event_id("Project Atlas budget", first.start.isoformat(), "Jane Doe")
    assert len(first.id) == 12 and int(first.id, 16) >= 0
    assert first.id != moved.id
    assert first.id != canceled.id


@pytest.mark.parametrize(
    ("label", "expected"),
    [
        ("November 15\u201321, 2026", DateRange(date(2026, 11, 15), date(2026, 11, 21))),
        ("November 29 \u2013 December 5, 2026", DateRange(date(2026, 11, 29), date(2026, 12, 5))),
        (
            "December 27, 2026 \u2013 January 2, 2027",
            DateRange(date(2026, 12, 27), date(2027, 1, 2)),
        ),
        ("Dec 27 - Jan 2, 2027", DateRange(date(2026, 12, 27), date(2027, 1, 2))),
        ("Nov 16-20, 2026", DateRange(date(2026, 11, 16), date(2026, 11, 20))),
    ],
)
def test_range_labels(label: str, expected: DateRange) -> None:
    assert parse_range_label(label) == expected


def test_range_label_rejects_plain_text() -> None:
    assert parse_range_label("Go to next week") is None


def test_view_label() -> None:
    label = "Change the view of your calendar, button. Work week selected"

    assert parse_view_label(label) == "work week"
    assert parse_view_label("Change the view of your calendar") is None
