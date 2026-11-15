from __future__ import annotations

from datetime import date

import pytest

from office_mcp.features.outlook.listing import has_load_more, parse_folder, parse_rows
from office_mcp.shared.node import Node

TODAY = date(2026, 10, 20)
NBSP = " "
ROWS = [
    "Today, Expanded",
    f"High priority, 3 unread messages,   3 messages, Alice Example, Bob Sample, Budget, Planning, and Review,"
    f"     6:53{NBSP}AM,        Hello team, the budget review moved to Thursday, ",
    f" Unread,    Carol Test, Weekly sync,  Meeting message,Tue 10/20/2026, 8:30{NBSP}AM "
    f"(30 minutes) No Conflicts   6:25{NBSP}AM,        Edit: moved to the afternoon, ",
    f"High priority,  Expanded,  3 messages,  Project Kickoff,  Meeting message,   3:34{NBSP}AM,"
    f"    Has files,     ",
    f"High priority,  Dan Demo, Canceled: Project Kickoff, Meeting message,  3:34{NBSP}AM,    Has files,"
    f"    The meeting is canceled, ",
    "  Erin Lee, Re: Project Kickoff,   Yesterday,        Thanks for the quick response",
    "  Carol Test, Re: Project Kickoff,   Yesterday,        We look forward to the meeting",
    f"    3 messages, Carol Test, Bob Sample, Quarterly report,     12:56{NBSP}AM,    Has files,"
    f"    Jane will follow up., ",
    "Last Month, Expanded",
    "    8 messages, Carol Test, Access request,   Replied  22/09/2026,"
    "        Travel update attached",
    "     Service Desk, Request received: REQ-100,     18/10/2026,        Hello Jane",
    "Load more conversations",
]


def rows() -> list[Node]:
    return [Node(role="AXRow", description=text, handle=index) for index, text in enumerate(ROWS)]


def test_parses_flags_counts_title_time_and_preview() -> None:
    parsed = [row.summary for row in parse_rows(rows(), TODAY)]

    first = parsed[0]
    assert first.flags == ("high priority",)
    assert (first.unread, first.unread_count, first.message_count) == (True, 3, 3)
    assert first.title == "Alice Example, Bob Sample, Budget, Planning, and Review"
    assert first.time_label == "6:53 AM"
    assert first.date == "2026-10-20"
    assert first.group == "Today"
    assert first.preview == "Hello team, the budget review moved to Thursday"


def test_meeting_rows_keep_meeting_details_out_of_the_title() -> None:
    meeting = parse_rows(rows(), TODAY)[1].summary

    assert meeting.title == "Carol Test, Weekly sync"
    assert meeting.meeting.startswith("Tue 10/20/2026, 8:30 AM (30 minutes) No Conflicts")
    assert meeting.time_label == "6:25 AM"
    assert meeting.unread and meeting.unread_count == 1


def test_expanded_conversation_folds_its_messages_into_a_thread() -> None:
    parsed = [row.summary for row in parse_rows(rows(), TODAY)]
    kickoff = parsed[2]

    assert "expanded" in kickoff.flags and "has attachments" in kickoff.flags
    assert kickoff.title == "Project Kickoff"
    assert len(kickoff.thread) == 3
    assert kickoff.thread[0].startswith("Dan Demo, Canceled: Project Kickoff | 3:34 AM")
    assert kickoff.thread[2].startswith("Carol Test, Re: Project Kickoff | Yesterday")
    assert [s.title for s in parsed[3:]] == [
        "Carol Test, Bob Sample, Quarterly report",
        "Carol Test, Access request",
        "Service Desk, Request received: REQ-100",
    ]


def test_status_prefix_and_attachment_flag_and_dates() -> None:
    parsed = [row.summary for row in parse_rows(rows(), TODAY)]
    report, replied, ticket = parsed[3], parsed[4], parsed[5]

    assert report.flags == ("has attachments",) and report.preview == "Jane will follow up."
    assert replied.flags == ("replied",)
    assert (replied.message_count, replied.date, replied.group) == (8, "2026-09-22", "Last Month")
    assert ticket.date == "2026-10-18" and ticket.preview == "Hello Jane"


def test_keys_are_stable_and_ignore_unread_state() -> None:
    before = parse_rows(rows(), TODAY)[0].summary.key
    read = [ROWS[0], ROWS[1].replace("3 unread messages,   ", ""), *ROWS[2:]]
    after = parse_rows([Node(role="AXRow", description=text) for text in read], TODAY)[
        0
    ].summary.key

    assert before == after


def test_load_more_row_is_found() -> None:
    assert has_load_more(rows()) is not None
    assert has_load_more(rows()[:-1]) is None


@pytest.mark.parametrize(
    ("description", "name", "unread"),
    [
        ("Inbox; 12 unread messages", "Inbox", 12),
        ("Drafts", "Drafts", 0),
        ("Projects; 1 unread message", "Projects", 1),
    ],
)
def test_folder_descriptions(description: str, name: str, unread: int) -> None:
    folder = parse_folder(description)
    assert folder is not None
    assert (folder.name, folder.unread) == (name, unread)


TWENTY_FOUR_HOUR_ROWS = [
    "Today, Expanded",
    " Unread,    Carol Test, Daily standup ,  Meeting message,Occurs every Tuesday, Thursday and "
    "Sunday effective 20 Oct 2026 until 1 Jan 2027 from 11:00 to 11:30. No Conflicts   14:05,"
    "        Dear team, this is the daily standup",
    " 1 unread message,   3 messages, Alice Example, Bob Sample, Platform features,     14:02,"
    "    Has files,    Hello team",
    "    2 messages, Dan Demo, Benchmarks,     08:09,        A gentle reminder",
    "Yesterday, Expanded",
    "    3 messages, Erin Lee, Planning,   Yesterday,        See you tomorrow",
]


def test_twenty_four_hour_times_are_parsed_as_today() -> None:
    nodes = [Node(role="AXRow", description=text) for text in TWENTY_FOUR_HOUR_ROWS]
    parsed = [row.summary for row in parse_rows(nodes, TODAY)]

    assert [(s.title, s.time_label, s.date) for s in parsed] == [
        ("Carol Test, Daily standup", "14:05", "2026-10-20"),
        ("Alice Example, Bob Sample, Platform features", "14:02", "2026-10-20"),
        ("Dan Demo, Benchmarks", "08:09", "2026-10-20"),
        ("Erin Lee, Planning", "Yesterday", "2026-10-19"),
    ]
    standup = parsed[0]
    assert standup.meeting.startswith("Occurs every Tuesday, Thursday and Sunday")
    assert standup.meeting.endswith("from 11:00 to 11:30. No Conflicts")
    assert standup.preview == "Dear team, this is the daily standup"
    assert parsed[1].flags == ("has attachments",) and parsed[1].unread_count == 1
