from __future__ import annotations

from dataclasses import replace
from datetime import date

import pytest
from fakes.p1_mail import OWNER_EMAIL, MailGateway, mail_rig, mail_row
from fakes.trees import folder_row, outlook_row, outlook_rows

from office_mcp.features.outlook.listing import parse_rows
from office_mcp.features.outlook.mail.followups import follow_ups, split_title, subject_key
from office_mcp.features.outlook.mail.service import sent_folder_name
from office_mcp.features.outlook.mail.triage import mail_needs_reply
from office_mcp.features.outlook.models import Folder, Summary
from office_mcp.features.outlook.service import MailView
from office_mcp.features.outlook.window import FocusedTab

TODAY = date(2026, 10, 20)


def summary(description: str, view: FocusedTab | None = FocusedTab.FOCUSED) -> Summary:
    parsed = parse_rows([outlook_row(description)], TODAY)[0].summary
    return replace(parsed, view=view.value if view else None)


ASK = summary(mail_row("Alice Example", "Project Atlas budget", "9:15 AM", "Can you check?"))
UNREAD_ASK = summary(mail_row("Omar Haddad", "Atlas review", "Yesterday", "Thoughts?", unread=True))
REPLIED = summary(mail_row("Alex Example", "Atlas timeline", "8:00 AM", "Done", flags="  Replied,"))
AUTOMATED = summary(
    mail_row("Atlas Notifications", "Build passed", "7:00 AM", "All green", unread=True)
)
NO_REPLY = summary(mail_row("no-reply@contoso.com", "Your receipt", "7:30 AM", "Thanks"))
MEETING = summary(
    " Unread,    Carol Test, Weekly sync,  Meeting message,Tue 10/20/2026, 8:30 AM "
    "(30 minutes) No Conflicts   6:25 AM,        Edit: moved"
)
OLD = summary(mail_row("Bob Sample", "Atlas kickoff", "01/10/2026", "Kickoff notes"))
OTHER_READ = summary(
    mail_row("Contoso News", "Weekly digest", "6:00 AM", "Top stories"), FocusedTab.OTHER
)
OTHER_UNREAD = summary(
    mail_row("Bob Sample", "Atlas invoice", "6:30 AM", "Please approve", unread=True),
    FocusedTab.OTHER,
)
FROM_OWNER = summary(mail_row(OWNER_EMAIL, "Atlas notes", "5:00 AM", "Notes for later"))


@pytest.mark.parametrize(
    ("row", "kept"),
    [
        (ASK, True),
        (UNREAD_ASK, True),
        (REPLIED, False),
        (AUTOMATED, False),
        (NO_REPLY, False),
        (MEETING, False),
        (OLD, False),
        (OTHER_READ, False),
        (OTHER_UNREAD, True),
        (FROM_OWNER, False),
    ],
)
def test_needs_reply_rules(row: Summary, kept: bool) -> None:
    found = mail_needs_reply([row], OWNER_EMAIL, TODAY, 7)

    assert bool(found) is kept


def test_needs_reply_orders_unread_first_then_newest_with_reasons() -> None:
    rows = [ASK, UNREAD_ASK, OTHER_UNREAD, REPLIED]

    found = mail_needs_reply(rows, OWNER_EMAIL, TODAY, 7)

    assert [(item.summary.title, item.reason) for item in found] == [
        ("Bob Sample, Atlas invoice", "unread, in the Other tab"),
        ("Omar Haddad, Atlas review", "unread"),
        ("Alice Example, Project Atlas budget", "read, no reply yet"),
    ]
    assert found[0].to_record()["reason"] == "unread, in the Other tab"
    assert found[0].to_record()["key"] == OTHER_UNREAD.key


def test_needs_reply_without_owner_email_keeps_owner_rows() -> None:
    assert mail_needs_reply([FROM_OWNER], None, TODAY, 7)


def test_needs_reply_service_reads_both_tabs_and_marks_nothing() -> None:
    gateway = MailGateway(
        rows=outlook_rows(
            mail_row("Alice Example", "Project Atlas budget", "9:15 AM", "Can you check?")
        )
    )
    gateway.tab_rows = {
        FocusedTab.FOCUSED: list(gateway.rows),
        FocusedTab.OTHER: [
            outlook_row(
                mail_row("Bob Sample", "Atlas invoice", "6:30 AM", "Approve?", unread=True),
                handle="other0",
            )
        ],
    }
    rig = mail_rig(gateway)

    found = rig.mail.needs_reply(7, "Inbox", MailView.BOTH, 10)

    assert [item.reason for item in found] == ["unread, in the Other tab", "read, no reply yet"]
    assert not any(call.startswith("select:row") for call in gateway.calls)
    assert gateway.tab is FocusedTab.FOCUSED


SENT_WAITING = summary(
    mail_row("Alice Example", "Project Atlas budget", "15/10/2026", "Numbers attached"), None
)
SENT_ANSWERED = summary(mail_row("Omar Haddad", "Atlas review", "14/10/2026", "Thoughts?"), None)
SENT_TOO_NEW = summary(mail_row("Alex Example", "Atlas notes", "9:00 AM", "Notes"), None)
SENT_TOO_OLD = summary(mail_row("Bob Sample", "Atlas kickoff", "01/09/2026", "Kickoff"), None)
SENT_OLDER_COPY = summary(
    mail_row("Alice Example", "Re: Project Atlas budget", "12/10/2026", "Earlier"), None
)
INBOX_REPLY = summary(mail_row("Omar Haddad", "RE: Atlas review", "16/10/2026", "Looks good"))
INBOX_BEFORE = summary(
    mail_row("Alice Example", "Project Atlas budget", "13/10/2026", "Before you sent it")
)
INBOX_OWN = summary(mail_row(OWNER_EMAIL, "Project Atlas budget", "17/10/2026", "Bump"))


def test_follow_ups_join_sent_and_inbox_by_subject() -> None:
    found = follow_ups(
        [SENT_WAITING, SENT_ANSWERED, SENT_TOO_NEW, SENT_TOO_OLD, SENT_OLDER_COPY],
        [INBOX_REPLY, INBOX_BEFORE, INBOX_OWN],
        OWNER_EMAIL,
        TODAY,
        2,
        14,
    )

    assert [item.to_record() for item in found] == [
        {
            "key": SENT_WAITING.key,
            "subject": "Project Atlas budget",
            "sent_date": "2026-10-15",
            "to_preview": "Alice Example",
            "days_waiting": 5,
        }
    ]


def test_subject_keys_ignore_reply_prefixes() -> None:
    assert subject_key("RE: Fw: Project Atlas budget") == subject_key("project atlas budget")
    assert split_title("Alice Example, Project Atlas budget") == (
        "Alice Example",
        "Project Atlas budget",
    )
    assert split_title("Untitled") == ("", "Untitled")


def test_sent_folder_name_prefers_what_outlook_shows() -> None:
    assert sent_folder_name([Folder("Inbox", 0), Folder("Sent", 0)]) == "Sent"
    assert sent_folder_name([Folder("Sent Items", 0)]) == "Sent Items"
    assert sent_folder_name([]) == "Sent Items"


def test_follow_ups_service_lists_sent_then_inbox() -> None:
    gateway = MailGateway(
        rows=outlook_rows(),
        folders=[folder_row("Inbox", handle="inbox"), folder_row("Sent", handle="sent")],
    )
    gateway.folder_rows_by_name = {
        "Sent": [outlook_row(mail_row("Alice Example", "Atlas plan", "15/10/2026", "Plan"))],
        "Inbox": [],
    }
    rig = mail_rig(gateway)

    found = rig.mail.follow_ups(2, 14, 10)

    assert [item.subject for item in found] == ["Atlas plan"]
    assert gateway.folder == "Inbox"
