from __future__ import annotations

import pytest
from fakes.p1_mail import FakeSuggestions, MailGateway, PaneSpec, mail_rig, mail_row, reading_pane
from fakes.trees import outlook_rows

from office_mcp.features.outlook.people.parse import (
    Contact,
    ContactSource,
    parse_suggestion,
    parse_suggestions,
    typed_rows,
)
from office_mcp.shared.errors import UiTimeoutError

TYPED_ROWS = [
    "Best Match text suggestion 1 of 4 highlighted. atlas budget",
    "Best Match people suggestion 2 of 4 Jane Doe jane@example.com",
    "Best Match people suggestion 3 of 4 Doe, John john.doe@contoso.com",
    "Best Match people suggestion 4 of 4 jane@example.com",
]


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        (
            "Best Match people suggestion 2 of 4 highlighted. Omar Haddad omar@example.com",
            Contact("Omar Haddad", "omar@example.com", ContactSource.SUGGESTION),
        ),
        (
            "Search Suggestion people suggestion 4 of 6 Alex Example <alex@example.com>",
            Contact("Alex Example", "alex@example.com", ContactSource.SUGGESTION),
        ),
        (
            "Best Match people suggestion 1 of 1 bob@contoso.com",
            Contact("bob@contoso.com", "bob@contoso.com", ContactSource.SUGGESTION),
        ),
        ("Best Match text suggestion 1 of 3 atlas", None),
        ("Best Match people suggestion 1 of 2 Jane Doe", None),
    ],
)
def test_parse_suggestion(text: str, expected: Contact | None) -> None:
    assert parse_suggestion(text) == expected


def test_parse_suggestions_drops_duplicates_and_text_rows() -> None:
    contacts = parse_suggestions(TYPED_ROWS)

    assert [contact.to_record() for contact in contacts] == [
        {"name": "Jane Doe", "email": "jane@example.com", "source": "outlook_suggestion"},
        {"name": "Doe, John", "email": "john.doe@contoso.com", "source": "outlook_suggestion"},
    ]


def test_rows_before_typing_are_not_trusted() -> None:
    history = ["Search Suggestion history suggestion 1 of 2 atlas"]

    assert typed_rows(history) is None
    assert typed_rows([*history, TYPED_ROWS[1]]) == [TYPED_ROWS[1]]


def test_lookup_reads_suggestions_and_clears_the_field() -> None:
    suggestions = FakeSuggestions(TYPED_ROWS)
    rig = mail_rig(MailGateway(rows=[]), suggestions=suggestions)

    contacts = rig.mail.people("jane", 5)

    assert [contact.email for contact in contacts] == ["jane@example.com", "john.doe@contoso.com"]
    assert suggestions.typed == ["jane"]
    assert suggestions.cleared == 1
    assert "run_search:jane:Current Mailbox" not in rig.gateway.calls


def test_lookup_merges_people_seen_while_reading() -> None:
    gateway = MailGateway(
        rows=outlook_rows(mail_row("Omar Haddad", "Atlas review", "8:40 AM", "See you"))
    )
    gateway.panes["row0"] = reading_pane(
        PaneSpec(subject="Atlas review", sender=("Omar Haddad", "omar@example.com"))
    )
    suggestions = FakeSuggestions([])
    rig = mail_rig(gateway, suggestions=suggestions)
    rig.mail.read("review", None)

    contacts = rig.mail.people("omar", 5)

    assert [contact.to_record() for contact in contacts] == [
        {"name": "Omar Haddad", "email": "omar@example.com", "source": "recent_recipient"}
    ]


def test_lookup_respects_the_limit() -> None:
    rig = mail_rig(MailGateway(rows=[]), suggestions=FakeSuggestions(TYPED_ROWS))

    assert len(rig.mail.people("jane", 1)) == 1


def test_field_is_cleared_even_when_typing_fails() -> None:
    suggestions = FakeSuggestions(fail=True)
    rig = mail_rig(MailGateway(rows=[]), suggestions=suggestions)

    with pytest.raises(UiTimeoutError):
        rig.mail.people("jane", 5)
    assert suggestions.cleared == 1
