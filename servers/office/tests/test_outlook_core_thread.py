from __future__ import annotations

from office_mcp.features.outlook.models import ThreadMessage
from office_mcp.features.outlook.thread import split_thread, top_text

TWELVE_HOUR = "\n".join(
    [
        "Thanks, that works for me.",
        "Jane",
        "From: Omar Haddad <omar@example.com>",
        "Sent: Sunday, November 15, 2026 2:13 PM",
        "To: Jane Doe <jane@example.com>",
        "Cc: Alex Example <alex@example.com>",
        "Subject: RE: Project Atlas budget",
        "Can we move the review to Thursday?",
        "From: Jane Doe <jane@example.com>",
        "Sent: Saturday, November 14, 2026 9:05 AM",
        "To: Omar Haddad <omar@example.com>",
        "Subject: Project Atlas budget",
        "Here is the budget.",
        "Numbers are attached.",
    ]
)

TWENTY_FOUR_HOUR = "\r\n".join(
    [
        "Noted.",
        "From:",
        "Alex Example",
        "Date: 15 November 2026 14:13",
        "To: Jane Doe",
        "Subject: Atlas timeline",
        "",
        "Timeline below.",
    ]
)


def test_splits_quoted_history_newest_first_with_twelve_hour_times() -> None:
    messages = split_thread(TWELVE_HOUR)

    assert messages == (
        ThreadMessage(
            sender="Omar Haddad <omar@example.com>",
            sent="Sunday, November 15, 2026 2:13 PM",
            to="Jane Doe <jane@example.com>",
            cc="Alex Example <alex@example.com>",
            subject="RE: Project Atlas budget",
            body="Can we move the review to Thursday?",
        ),
        ThreadMessage(
            sender="Jane Doe <jane@example.com>",
            sent="Saturday, November 14, 2026 9:05 AM",
            to="Omar Haddad <omar@example.com>",
            cc="",
            subject="Project Atlas budget",
            body="Here is the budget.\nNumbers are attached.",
        ),
    )
    assert top_text(TWELVE_HOUR) == "Thanks, that works for me.\nJane"


def test_split_label_lines_date_header_and_twenty_four_hour_times() -> None:
    (message,) = split_thread(TWENTY_FOUR_HOUR)

    assert message.sender == "Alex Example"
    assert message.sent == "15 November 2026 14:13"
    assert message.subject == "Atlas timeline"
    assert message.body == "Timeline below."
    assert top_text(TWENTY_FOUR_HOUR) == "Noted."


def test_a_from_line_without_a_date_nearby_is_body_text() -> None:
    body = "From: the team\nWe shipped it.\nMore text\nAnd more\nSent: later"

    assert split_thread(body) == ()
    assert top_text(body) == body


def test_limit_caps_the_number_of_messages() -> None:
    assert len(split_thread(TWELVE_HOUR, limit=1)) == 1
    assert split_thread(TWELVE_HOUR, limit=0) == ()
