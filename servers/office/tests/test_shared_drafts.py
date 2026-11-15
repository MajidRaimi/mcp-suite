from __future__ import annotations

import pytest

from office_mcp.shared.drafts import (
    Channel,
    Draft,
    Fix,
    NoTextGuard,
    Review,
    Severity,
    Violation,
    canonical_text,
    written_exactly,
)
from office_mcp.shared.errors import (
    AmbiguousTargetError,
    OfficeError,
    PersonaViolationError,
    TextEntryError,
    error_text,
)

EM_DASH = "\u2014"
EN_DASH = "\u2013"


def test_canonical_text_unifies_line_endings_and_trims() -> None:
    raw = "\r\n\nHi Jane,  \r\nthe plan\t\rworks\n\n\n"

    assert canonical_text(raw) == "Hi Jane,\nthe plan\nworks"


def test_canonical_text_keeps_unicode_as_is() -> None:
    assert canonical_text("caf\u00e9") != canonical_text("cafe\u0301")


def test_written_exactly_tolerates_crlf_and_trailing_spaces() -> None:
    assert written_exactly("Hi Jane,\nthe plan works", "Hi Jane,  \r\nthe plan works\n")


def test_written_exactly_detects_a_dash_substitution() -> None:
    assert not written_exactly("the plan - then the budget", f"the plan {EM_DASH} then the budget")
    assert not written_exactly("2024-2025", f"2024{EN_DASH}2025")


def test_written_exactly_detects_a_joined_closing_line() -> None:
    assert not written_exactly("Thanks\n\nBest regards,", "Thanks Best regards,")


def test_no_text_guard_reviews_clean_and_verifies_exactness() -> None:
    guard = NoTextGuard()
    draft = Draft(Channel.EMAIL, "Hi Jane,\nall set.", subject="Project Atlas budget")

    review = guard.require_clean(draft)

    assert review.ok and review.persona_version == "none"
    assert review.draft is draft
    guard.verify_written(draft, "Hi Jane,\r\nall set.\n", "Project Atlas budget")
    with pytest.raises(TextEntryError, match=r"line 2: expected \"all set.\", found \"all set!\""):
        guard.verify_written(draft, "Hi Jane,\nall set!", "Project Atlas budget")
    with pytest.raises(TextEntryError, match="subject"):
        guard.verify_written(draft, "Hi Jane,\nall set.", "RE: Project Atlas budget")


def test_text_entry_diff_shows_dashes_escaped() -> None:
    draft = Draft(Channel.TEAMS, "the plan - then")

    with pytest.raises(TextEntryError) as caught:
        NoTextGuard().verify_written(draft, f"the plan {EM_DASH} then", "")

    assert "[U+2014]" in str(caught.value)
    assert EM_DASH not in str(caught.value)


def test_review_record_and_properties() -> None:
    draft = Draft(Channel.EMAIL, "Hi\n\nBest regards,")
    review = Review(
        draft=draft,
        violations=(
            Violation("length", "over 180 words", Severity.WARNING),
            Violation("characters", f"forbidden {EM_DASH}", Severity.ERROR, excerpt=EM_DASH),
        ),
        fixes=(Fix("signoff.appended", "added Best regards,"),),
        persona_version="abc123",
    )

    record = review.to_record()

    assert not review.ok
    assert review.warnings == ("over 180 words",)
    assert record["final_text"] == "Hi\n\nBest regards,"
    assert record["fixes"] == [{"rule": "signoff.appended", "detail": "added Best regards,"}]
    assert record["violations"][1]["excerpt"] == "[U+2014]"
    assert record["persona_version"] == "abc123"


def test_persona_violation_error_text_format() -> None:
    error = PersonaViolationError(
        [
            Violation(
                "characters",
                f'forbidden character U+2014 EM DASH in body: "...the plan {EM_DASH} then..."',
                Severity.ERROR,
                hint="Use a comma, period or colon.",
            ),
            Violation(
                "signoff.trailing",
                'nothing may follow "Best regards," (the signature card is attached '
                'automatically). Found "Alex".',
                Severity.ERROR,
            ),
            Violation("length", "long", Severity.WARNING),
        ]
    )

    assert str(error) == (
        "Not sent: the text breaks 2 persona rules. Fix them and call the tool again.\n"
        '1. characters: forbidden character U+2014 EM DASH in body: "...the plan [U+2014] '
        'then...". Use a comma, period or colon.\n'
        '2. signoff.trailing: nothing may follow "Best regards," (the signature card is '
        'attached automatically). Found "Alex".'
    )
    assert len(error.violations) == 3


def test_error_text_joins_notes_and_ambiguous_lists_candidates() -> None:
    error = AmbiguousTargetError("two rows match", ["a1b2c3d4e5f6", "f6e5d4c3b2a1"])
    error.add_note("cleanup close menu failed: stuck")

    assert error.candidates == ("a1b2c3d4e5f6", "f6e5d4c3b2a1")
    assert error_text(error) == (
        "two rows match Candidates: a1b2c3d4e5f6; f6e5d4c3b2a1\ncleanup close menu failed: stuck"
    )
    assert error_text(OfficeError("plain")) == "plain"
