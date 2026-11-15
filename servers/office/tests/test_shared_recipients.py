from __future__ import annotations

import pytest

from office_mcp.shared.errors import RecipientNotAllowedError, SettingsError
from office_mcp.shared.recipients import (
    RecipientGuard,
    is_email,
    normalize_email,
    parse_allowlist,
)

SELF_THREAD = "48:notes"


def test_parse_allowlist_accepts_the_four_forms() -> None:
    allowlist = parse_allowlist(
        "Jane@Example.com, @contoso.com, teams:self, teams:19:abc_def@unq.gbl.spaces"
    )

    assert allowlist is not None
    assert allowlist.emails == frozenset({"jane@example.com"})
    assert allowlist.domains == frozenset({"contoso.com"})
    assert allowlist.teams_threads == frozenset({"19:abc_def@unq.gbl.spaces"})
    assert allowlist.teams_self
    assert allowlist.entries() == (
        "jane@example.com",
        "@contoso.com",
        "teams:self",
        "teams:19:abc_def@unq.gbl.spaces",
    )


def test_unset_allowlist_is_none() -> None:
    assert parse_allowlist(None) is None
    assert parse_allowlist("  ") is None


@pytest.mark.parametrize("raw", ["jane", "@", "@bad domain", "teams:", "slack:general", "a@b"])
def test_invalid_entries_raise(raw: str) -> None:
    with pytest.raises(SettingsError):
        parse_allowlist(raw)


def test_inactive_guard_allows_everything() -> None:
    guard = RecipientGuard(None, SELF_THREAD)

    guard.require_emails(["omar@example.org"])
    guard.require_thread("19:anything@thread.v2")

    assert not guard.active
    assert guard.describe() is None


def test_email_and_domain_matching() -> None:
    guard = RecipientGuard(parse_allowlist("jane@example.com,@contoso.com"), SELF_THREAD)

    guard.require_emails([" JANE@example.com ", "bob.sample@contoso.com"])
    with pytest.raises(RecipientNotAllowedError, match=r"omar@example\.org"):
        guard.require_emails(["jane@example.com", "omar@example.org"])
    with pytest.raises(RecipientNotAllowedError, match=r"alex@sub\.contoso\.com"):
        guard.require_emails(["alex@sub.contoso.com"])
    assert guard.describe() == ["jane@example.com", "@contoso.com"]


def test_teams_self_allows_only_the_configured_thread() -> None:
    guard = RecipientGuard(parse_allowlist("teams:self"), "19:me@unq.gbl.spaces")

    guard.require_thread("19:me@unq.gbl.spaces")
    with pytest.raises(RecipientNotAllowedError, match="teams:48:notes"):
        guard.require_thread("48:notes")
    with pytest.raises(RecipientNotAllowedError):
        guard.require_emails(["jane@example.com"])


def test_explicit_thread_entries() -> None:
    guard = RecipientGuard(parse_allowlist("teams:48:notes"), "")

    guard.require_thread("48:notes")
    with pytest.raises(RecipientNotAllowedError, match="Nothing was sent"):
        guard.require_thread("19:other@thread.v2")


def test_is_email_and_normalize() -> None:
    assert is_email("nobody@example.invalid")
    assert is_email("jane.doe+tag@mail.example.com")
    assert not is_email("jane@@example.com")
    assert not is_email("Jane Doe <jane@example.com>")
    assert normalize_email(" Jane@Example.COM ") == "jane@example.com"
