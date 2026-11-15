from __future__ import annotations

import pytest
from fakes.clock import FakeClock
from fakes.drivers import FAKE_TODAY
from fakes.p1_mail import (
    OWNER_APP,
    MailGateway,
    PaneSpec,
    attachment_group,
    mail_rig,
    mail_row,
    reading_pane,
)
from fakes.trees import folder_row, load_more_row, outlook_row, outlook_rows

from office_mcp.features.outlook.listing import parse_rows
from office_mcp.features.outlook.reading import parse_email
from office_mcp.features.outlook.service import MailView, clock_minutes, merge_newest_first
from office_mcp.features.outlook.window import FocusedTab
from office_mcp.shared.errors import ActionNotVerifiedError, TargetError
from office_mcp.shared.runner import Deadline

TODAY_ROW = mail_row("Alice Example", "Project Atlas budget", "9:15 AM", "Numbers attached")
UNREAD_ROW = mail_row("Omar Haddad", "Atlas review", "8:40 AM", "See you at ten", unread=True)
OTHER_ROW = mail_row("Bob Sample", "Atlas newsletter", "10:05 AM", "This week in Atlas")
QUOTED_BODY = (
    "Thanks, see the numbers below.",
    "From: Omar Haddad <omar@example.com>",
    "Sent: Monday, October 19, 2026 4:10 PM",
    "To: Alice Example <alice@example.com>",
    "Subject: Atlas review",
    "Can you send the numbers?",
)


def gateway_with(*descriptions: str) -> MailGateway:
    return MailGateway(rows=outlook_rows(*descriptions))


def test_read_restores_unread_once_and_deselects() -> None:
    gateway = gateway_with(TODAY_ROW, UNREAD_ROW)
    gateway.panes["row1"] = reading_pane(PaneSpec(subject="Atlas review"))
    rig = mail_rig(gateway)

    email = rig.mail.read("Atlas review", None)

    assert email.subject == "Atlas review"
    assert email.summary is not None and email.summary.unread
    assert rig.marker.presses == 1
    assert gateway.is_unread("row1")
    assert "deselect:row1" in gateway.calls
    assert gateway.selected == set()
    assert email.notes == ()
    assert rig.rig.workspace.activations[0] == rig.rig.app.app.require_pid()
    assert rig.rig.workspace.frontmost == OWNER_APP


def test_read_of_a_read_conversation_never_marks_or_comes_forward() -> None:
    gateway = gateway_with(TODAY_ROW, UNREAD_ROW)
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Project Atlas budget"))
    rig = mail_rig(gateway)

    rig.mail.read("budget", None)

    assert rig.marker.presses == 0
    assert rig.rig.workspace.activations == []
    assert "deselect:row0" in gateway.calls


def test_keep_unread_false_leaves_it_read() -> None:
    gateway = gateway_with(UNREAD_ROW)
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Atlas review"))
    rig = mail_rig(gateway)

    rig.mail.read("review", None, keep_unread=False)

    assert rig.marker.presses == 0
    assert not gateway.is_unread("row0")
    assert gateway.selected == set()


def test_unread_restore_is_corrected_once_then_reported() -> None:
    gateway = gateway_with(UNREAD_ROW)
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Atlas review"))
    rig = mail_rig(gateway, ignored_unread=1)

    email = rig.mail.read("review", None)

    assert rig.marker.presses == 2
    assert gateway.is_unread("row0")
    assert email.notes == ()


def test_a_mail_that_turns_read_when_deselected_is_marked_unread_again() -> None:
    gateway = gateway_with(UNREAD_ROW)
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Atlas review"))
    gateway.read_on_deselect = True
    rig = mail_rig(gateway)

    email = rig.mail.read("review", None)

    assert rig.marker.presses == 1
    assert gateway.is_unread("row0")
    assert gateway.selected == set()
    assert email.notes == ()


EXPANDED_HEADER = "1 unread message, Expanded,  2 messages,  Atlas review,     8:40 AM,         "


def test_a_conversation_expanded_by_the_read_is_collapsed_again() -> None:
    gateway = gateway_with(UNREAD_ROW)
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Atlas review"))
    gateway.expand_on_select["row0"] = EXPANDED_HEADER
    rig = mail_rig(gateway)

    email = rig.mail.read("review", None)

    assert "collapse" in gateway.calls
    assert gateway.collapsed == {}
    assert gateway.is_unread("row0")
    assert email.notes == ()


def test_a_conversation_that_stays_expanded_is_reported() -> None:
    gateway = gateway_with(OTHER_ROW)
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Atlas newsletter"))
    gateway.expand_on_select["row0"] = "Expanded,  2 messages,  Atlas newsletter,     10:05 AM,  "
    gateway.collapse_ignored = True
    rig = mail_rig(gateway)

    email = rig.mail.read("newsletter", None)

    assert email.notes == ("'Bob Sample, Atlas newsletter' was left expanded in the message list",)


def test_unread_restore_failure_becomes_a_note() -> None:
    gateway = gateway_with(UNREAD_ROW)
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Atlas review"))
    rig = mail_rig(gateway, ignored_unread=5)

    email = rig.mail.read("review", None)

    assert rig.marker.presses == 2
    assert email.notes == (
        "could not mark 'Omar Haddad, Atlas review' unread again; it is now read",
    )


def test_latest_message_restores_unread_and_deselects() -> None:
    gateway = gateway_with(UNREAD_ROW)
    gateway.panes["row0"] = reading_pane(
        PaneSpec(
            subject="Atlas review",
            sender=("Omar Haddad", "omar@example.com"),
            cc=(("Alex Example", "alex@example.com"),),
            body=QUOTED_BODY,
        )
    )
    rig = mail_rig(gateway)
    key = rig.mail.conversations("Inbox", False, None, 0)[0].key

    latest = rig.mail.latest_message(key, None)

    assert latest.key == key
    assert latest.sender is not None and latest.sender.email == "omar@example.com"
    assert [person.email for person in latest.to] == ["jane@example.com"]
    assert [person.email for person in latest.cc] == ["alex@example.com"]
    assert latest.body_top == "Thanks, see the numbers below."
    assert rig.marker.presses == 1
    assert gateway.is_unread("row0")
    assert gateway.selected == set()


def test_latest_message_takes_keys_only() -> None:
    rig = mail_rig(gateway_with(UNREAD_ROW))

    with pytest.raises(TargetError, match="use the key"):
        rig.mail.latest_message("Atlas review", None)
    assert rig.gateway.selected == set()


def test_latest_message_raises_when_unread_cannot_be_restored() -> None:
    gateway = gateway_with(UNREAD_ROW)
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Atlas review"))
    rig = mail_rig(gateway, ignored_unread=5)
    key = rig.mail.conversations("Inbox", False, None, 0)[0].key

    with pytest.raises(ActionNotVerifiedError, match="unread again"):
        rig.mail.latest_message(key, None)


def test_read_returns_thread_messages_and_attachments() -> None:
    gateway = gateway_with(TODAY_ROW)
    gateway.panes["row0"] = reading_pane(
        PaneSpec(
            subject="Project Atlas budget",
            body=QUOTED_BODY,
            attachments=(attachment_group("budget.xlsx", "12.5 KB"),),
        )
    )
    rig = mail_rig(gateway)

    record = rig.mail.read("budget", None, max_messages=5).to_record()

    first, second = record["messages"]
    assert first["from"] == "Alice Example <alice@example.com>"
    assert first["to"] == "Jane Doe <jane@example.com>"
    assert first["body"] == "Thanks, see the numbers below."
    assert second["from"] == "Omar Haddad <omar@example.com>"
    assert second["subject"] == "Atlas review"
    assert record["attachments"] == [{"name": "budget.xlsx", "size": "12.5 KB"}]


def test_max_messages_one_keeps_only_the_newest() -> None:
    pane = reading_pane(PaneSpec(subject="Atlas review", body=QUOTED_BODY))

    email = parse_email(pane, max_messages=1)

    assert email is not None
    assert len(email.messages) == 1


def test_view_both_merges_tabs_and_restores_the_tab() -> None:
    gateway = gateway_with(TODAY_ROW, UNREAD_ROW)
    gateway.tab_rows = {
        FocusedTab.FOCUSED: outlook_rows(TODAY_ROW, UNREAD_ROW),
        FocusedTab.OTHER: [outlook_row(OTHER_ROW, handle="other0")],
    }
    rig = mail_rig(gateway)

    summaries = rig.mail.conversations("Inbox", False, None, 0, view=MailView.BOTH)

    assert [summary.title for summary in summaries] == [
        "Bob Sample, Atlas newsletter",
        "Alice Example, Project Atlas budget",
        "Omar Haddad, Atlas review",
    ]
    assert [summary.view for summary in summaries] == ["other", "focused", "focused"]
    assert gateway.tab is FocusedTab.FOCUSED
    assert gateway.visible_calls().count("switch_tab:other") == 1
    assert gateway.visible_calls()[-1] == "switch_tab:focused"


def test_view_both_without_tabs_reads_once() -> None:
    gateway = MailGateway(rows=outlook_rows(TODAY_ROW), tab=None)
    rig = mail_rig(gateway)

    summaries = rig.mail.conversations("Inbox", False, None, 0, view=MailView.BOTH)

    assert len(summaries) == 1
    assert not any(call.startswith("switch_tab") for call in gateway.calls)


def test_budget_expiry_stops_load_more_before_any_press() -> None:
    gateway = MailGateway(rows=[*outlook_rows(TODAY_ROW), load_more_row()])
    gateway.pages = [outlook_rows(OTHER_ROW)]
    rig = mail_rig(gateway)
    clock = FakeClock(start=50.0)

    summaries = rig.mail.conversations(
        "Inbox", False, None, 0, load_more=5, deadline=Deadline(clock, 10.0)
    )

    assert len(summaries) == 1
    assert "press:load-more" not in gateway.calls


def test_budget_expiry_stops_load_more_midway() -> None:
    first_page = [outlook_row(OTHER_ROW, handle="page1"), load_more_row()]
    second_page = [outlook_row(UNREAD_ROW, handle="page2")]
    gateway = MailGateway(rows=[*outlook_rows(TODAY_ROW), load_more_row()])
    gateway.pages = [first_page, second_page]
    clock = FakeClock()
    gateway.on_press = lambda _node: clock.advance(30.0)
    rig = mail_rig(gateway)

    summaries = rig.mail.conversations(
        "Inbox", False, None, 0, load_more=5, deadline=Deadline(clock, 20.0)
    )

    assert gateway.calls.count("press:load-more") == 1
    assert len(summaries) == 2


def test_days_loads_pages_until_covered() -> None:
    older = mail_row("Alex Example", "Atlas kickoff", "12/10/2026", "Kickoff notes")
    gateway = MailGateway(rows=[*outlook_rows(TODAY_ROW), load_more_row()])
    gateway.pages = [[outlook_row(older, handle="old"), load_more_row()], outlook_rows(OTHER_ROW)]
    rig = mail_rig(gateway)

    summaries = rig.mail.conversations("Inbox", False, 3, 0)

    assert gateway.calls.count("press:load-more") == 1
    assert [summary.title for summary in summaries] == ["Alice Example, Project Atlas budget"]


def test_folders_skip_favorites() -> None:
    gateway = MailGateway(
        folders=[
            folder_row("Favorites", level=0),
            folder_row("Inbox; 3 unread messages", level=1, handle="fav-inbox"),
            folder_row("jane@example.com", level=0),
            folder_row("Inbox; 3 unread messages", level=1, handle="inbox"),
            folder_row("Drafts; 2 messages", level=1),
        ]
    )
    rig = mail_rig(gateway)

    folders = rig.mail.folders()

    assert [(folder.path, folder.unread, folder.total) for folder in folders] == [
        ("jane@example.com", 0, None),
        ("jane@example.com/Inbox", 3, None),
        ("jane@example.com/Drafts", 0, 2),
    ]


def test_merge_sorts_by_date_then_clock() -> None:
    rows = parse_rows(outlook_rows(UNREAD_ROW, TODAY_ROW), FAKE_TODAY)
    merged = merge_newest_first([row.summary for row in rows])

    assert [summary.time_label for summary in merged] == ["9:15 AM", "8:40 AM"]
    assert clock_minutes("13:05") == 785
    assert clock_minutes("12:30 AM") == 30
    assert clock_minutes("Yesterday") == -1


def test_listed_keys_remember_their_tab() -> None:
    gateway = gateway_with(TODAY_ROW)
    gateway.tab_rows = {
        FocusedTab.FOCUSED: outlook_rows(TODAY_ROW),
        FocusedTab.OTHER: [outlook_row(OTHER_ROW, handle="other0")],
    }
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Project Atlas budget"))
    rig = mail_rig(gateway)
    key = rig.mail.conversations("Inbox", False, None, 0)[0].key
    gateway.switch_tab(FocusedTab.OTHER)

    email = rig.mail.read(key, None)

    assert email.summary is not None and email.summary.view == "focused"
    assert gateway.tab is FocusedTab.OTHER
