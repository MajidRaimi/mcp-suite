from __future__ import annotations

from dataclasses import replace

import pytest
from fakes.clock import FakeClock
from fakes.drivers import FAKE_NOW, FAKE_TODAY
from fakes.p3_compose import OWNER, latest_message, sent_row

from office_mcp.features.outlook.compose.gateway import (
    RecipientField,
    RecipientRead,
    RecipientToken,
    ReplyKind,
)
from office_mcp.features.outlook.compose.guard import SendGuard, send_fingerprint
from office_mcp.features.outlook.compose.recipients import (
    parse_token,
    pending_text,
    reply_recipients,
    require_distinct,
    require_tokens,
    token_addresses,
    validate_addresses,
)
from office_mcp.features.outlook.compose.verify import (
    base_subject,
    body_is_free,
    display_subject,
    matching_rows,
    require_subject,
    safe_alert_button,
    sent_recently,
    written_part,
)
from office_mcp.features.outlook.listing import parse_rows
from office_mcp.features.outlook.models import Person
from office_mcp.shared.errors import ComposeError, DuplicateSendError, TextEntryError


@pytest.mark.parametrize(
    ("subject", "base"),
    [
        ("RE: Project Atlas budget", "project atlas budget"),
        ("Re: RE: FW: Project Atlas budget", "project atlas budget"),
        ("Fwd:Project  Atlas budget", "project atlas budget"),
        ("Project Atlas budget", "project atlas budget"),
    ],
)
def test_base_subject_strips_reply_and_forward_prefixes(subject: str, base: str) -> None:
    assert base_subject(subject) == base


def test_display_subject_uses_one_prefix() -> None:
    assert display_subject(ReplyKind.REPLY, "RE: Budget") == "RE: Budget"
    assert display_subject(ReplyKind.FORWARD, "RE: Budget") == "FW: Budget"


def test_require_subject_checks_prefix_kind_and_base() -> None:
    require_subject(ReplyKind.REPLY, "Budget", "Re: Budget")
    require_subject(ReplyKind.FORWARD, "RE: Budget", "FW: Budget")
    with pytest.raises(ComposeError):
        require_subject(ReplyKind.REPLY, "Budget", "FW: Budget")
    with pytest.raises(ComposeError):
        require_subject(ReplyKind.REPLY_ALL, "Budget", "RE: Other")
    with pytest.raises(ComposeError):
        require_subject(ReplyKind.REPLY, "Budget", "Budget")


@pytest.mark.parametrize(
    ("baseline", "free"),
    [
        ("", True),
        ("\nJane Doe", True),
        ("￼\nJane Doe", True),
        ("  \n\nFrom: Omar", True),
        ("Draft text\nJane Doe", False),
    ],
)
def test_body_is_free_only_with_an_empty_first_line(baseline: str, free: bool) -> None:
    assert body_is_free(baseline) is free


def test_written_part_removes_the_unchanged_signature() -> None:
    assert written_part("Hello\nJane Doe", "\nJane Doe") == "Hello"
    assert written_part("Hello", "") == "Hello"
    with pytest.raises(TextEntryError):
        written_part("Hello\nJane", "\nJane Doe")


def test_safe_alert_button_never_picks_send_anyway() -> None:
    assert safe_alert_button(("Send Anyway", "Cancel")) == "Cancel"
    assert safe_alert_button(("Send Anyway", "Don't Send")) == "Don't Send"
    assert safe_alert_button(("Continue Editing", "Discard")) == "Continue Editing"
    assert safe_alert_button(("Send Anyway", "OK")) is None


@pytest.mark.parametrize(
    ("label", "recent"),
    [("9:25 AM", True), ("9:30 AM", True), ("9:19 AM", False), ("8:00 AM", False)],
)
def test_sent_recently_uses_the_row_time(label: str, recent: bool) -> None:
    summary = parse_rows([sent_row("Budget", label)], FAKE_TODAY)[0].summary
    assert sent_recently(summary, FAKE_NOW) is recent


def test_sent_recently_handles_24_hour_and_other_days() -> None:
    today = parse_rows([sent_row("Budget", "09:26")], FAKE_TODAY)[0].summary
    older = parse_rows([sent_row("Budget", "Yesterday")], FAKE_TODAY)[0].summary
    assert sent_recently(today, FAKE_NOW) is True
    assert sent_recently(older, FAKE_NOW) is False


def test_matching_rows_needs_the_subject_at_the_end_of_the_title() -> None:
    rows = parse_rows(
        [sent_row("Project Atlas budget"), sent_row("Project Atlas budget review")], FAKE_TODAY
    )
    found = matching_rows(rows, "project atlas budget")
    assert [row.summary.title for row in found] == ["Omar Haddad, Project Atlas budget"]


def test_validate_addresses_limits_and_names() -> None:
    assert validate_addresses(RecipientField.TO, [" bob@example.com "], 1) == ("bob@example.com",)
    with pytest.raises(ComposeError, match="outlook_people"):
        validate_addresses(RecipientField.TO, ["Bob Sample"], 1)
    with pytest.raises(ComposeError, match="1 to 20"):
        validate_addresses(RecipientField.TO, [], 1)
    with pytest.raises(ComposeError):
        validate_addresses(RecipientField.CC, [f"u{n}@example.com" for n in range(21)], 0)


def test_require_distinct_across_fields() -> None:
    require_distinct(["a@example.com", "b@example.com"])
    with pytest.raises(ComposeError, match="once"):
        require_distinct(["a@example.com", "A@Example.com"])


def test_reply_goes_to_the_sender() -> None:
    latest = latest_message("0123456789ab")
    assert reply_recipients(latest, OWNER, reply_all=False) == ("omar@example.com",)


def test_reply_to_own_message_goes_to_its_recipients() -> None:
    latest = latest_message(
        "0123456789ab",
        sender=Person("Jane Doe", OWNER),
        to=(Person("Bob Sample", "bob@example.com"),),
    )
    assert reply_recipients(latest, OWNER, reply_all=False) == ("bob@example.com",)


def test_reply_all_needs_the_owner_address() -> None:
    with pytest.raises(ComposeError, match="OFFICE_MCP_OWNER_EMAIL"):
        reply_recipients(latest_message("0123456789ab"), None, reply_all=True)


def test_unreadable_sender_is_refused() -> None:
    latest = replace(latest_message("0123456789ab"), sender=Person("Omar Haddad", ""))
    with pytest.raises(ComposeError, match="sender address"):
        reply_recipients(latest, OWNER, reply_all=False)


def test_parse_token_reads_an_embedded_address() -> None:
    assert parse_token("Bob Sample, bob@Example.com").address == "bob@example.com"
    assert parse_token("Bob Sample", ("", "<bob@example.com>")).address == "bob@example.com"
    assert parse_token("Bob Sample").address is None


def test_pending_text_ignores_separators_and_token_marks() -> None:
    assert pending_text("￼; ￼,  ") == ""
    assert pending_text("￼ bob@exa") == "bob@exa"


def test_token_addresses_and_counts() -> None:
    read = RecipientRead(tokens=(RecipientToken("Bob", "bob@example.com"),), pending="")
    found = token_addresses(RecipientField.TO, read)
    require_tokens("To", found, ["BOB@example.com"])
    with pytest.raises(ComposeError, match="differ"):
        require_tokens("To", found, ["eve@example.com"])
    with pytest.raises(ComposeError, match="shows 1"):
        require_tokens("To", found, ["bob@example.com", "eve@example.com"])


def test_send_guard_window_and_last() -> None:
    clock = FakeClock(100.0)
    guard = SendGuard(clock)
    first = send_fingerprint("send", ["b@example.com"], "Hi", "Text\r\n")
    same = send_fingerprint("send", ["B@example.com"], "hi", "Text")
    assert first == same
    entry = guard.record(first, "Hi")
    with pytest.raises(DuplicateSendError):
        guard.check(same)
    assert guard.last(30) is entry
    clock.advance(31)
    assert guard.last(30) is None
    clock.advance(90)
    guard.check(same)
    guard.forget(entry)
    assert guard.last(1000) is None
