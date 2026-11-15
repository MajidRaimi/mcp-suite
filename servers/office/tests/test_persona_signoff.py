from __future__ import annotations

from pathlib import Path

import pytest
from fakes.p7_persona import guard_for, review, rules_hit

from office_mcp.features.persona.rules.signoff import is_sign_off_like
from office_mcp.shared.drafts import Channel

ARABIC_BODY = "شكرا على التحديث، سأرسل الملف غدا"
ARABIC_CONTEXT = "هل انتهيت من ميزانية المشروع"
ARABIC_CLOSING = "مع التحية"
DEFAULT_ONLY = """+++
[email]
sign_off = { default = "Best regards," }
sign_off_rule = "required"
nothing_after_sign_off = true
+++
"""


def test_closing_is_appended_when_missing(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "Hi Jane,\n\nThe budget is approved.")

    assert result.ok
    assert result.draft.body == "Hi Jane,\n\nThe budget is approved.\n\nBest regards,"
    assert [fix.rule for fix in result.fixes] == ["signoff.appended"]


def test_closing_is_not_appended_twice(tmp_path: Path) -> None:
    body = "Hi Jane,\n\nThe budget is approved.\n\nBest regards,"

    result = review(guard_for(tmp_path), body)

    assert result.ok
    assert result.draft.body == body
    assert result.fixes == ()


def test_review_of_the_final_text_is_stable(tmp_path: Path) -> None:
    guard = guard_for(tmp_path)
    first = review(guard, "The budget is approved.")

    second = review(guard, first.draft.body)

    assert second.draft.body == first.draft.body
    assert second.fixes == ()


@pytest.mark.parametrize(
    "closing", ["Thanks,", "Kind regards", "Cheers!", "Best wishes,", "Regards, Alex"]
)
def test_wrong_closing_is_refused(tmp_path: Path, closing: str) -> None:
    result = review(guard_for(tmp_path), f"The budget is approved.\n\n{closing}")

    assert rules_hit(result) == ["signoff.wrong"]
    assert result.fixes == ()
    assert result.draft.body.endswith(closing)


def test_name_after_closing_is_refused(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "The budget is approved.\n\nBest regards,\nAlex")

    assert rules_hit(result) == ["signoff.trailing"]
    assert result.violations[0].message == (
        'nothing may follow "Best regards," (the signature card is attached automatically). '
        'Found "Alex".'
    )
    assert result.fixes == ()


def test_arabic_body_uses_the_arabic_closing(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), ARABIC_BODY, context=ARABIC_CONTEXT)

    assert result.ok
    assert result.draft.body.endswith(f"\n\n{ARABIC_CLOSING}")


def test_arabic_body_falls_back_to_the_default_closing(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path, DEFAULT_ONLY), ARABIC_BODY)

    assert result.draft.body.endswith("\n\nBest regards,")


def test_arabic_sign_off_like_line_is_refused(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), f"{ARABIC_BODY}\n\nوشكرا\u064b", context=ARABIC_CONTEXT)

    assert rules_hit(result) == ["signoff.wrong"]


def test_other_account_skips_the_closing(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "The budget is approved.", account="alex@contoso.com")

    assert result.ok
    assert result.fixes == ()
    assert result.draft.body == "The budget is approved."


def test_matching_account_is_case_insensitive(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "Approved.", account="Alex@Example.COM")

    assert [fix.rule for fix in result.fixes] == ["signoff.appended"]


def test_unknown_account_still_applies_the_rule(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "Approved.", account="")

    assert [fix.rule for fix in result.fixes] == ["signoff.appended"]


@pytest.mark.parametrize("last", ["Thanks", "Thanks!", "Best regards,", "شكرا"])
def test_teams_closing_is_forbidden(tmp_path: Path, last: str) -> None:
    result = review(guard_for(tmp_path), f"Done, the file is uploaded.\n{last}", Channel.TEAMS)

    assert "signoff.forbidden" in rules_hit(result)


def test_teams_without_closing_is_fine(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "Done, the file is uploaded.", Channel.TEAMS)

    assert result.ok
    assert result.fixes == ()


@pytest.mark.parametrize("body", ["Thanks!", "Got it.\nThanks for the update", "Best option is B."])
def test_teams_content_lines_are_not_taken_for_closings(tmp_path: Path, body: str) -> None:
    result = review(guard_for(tmp_path), body, Channel.TEAMS)

    assert "signoff.forbidden" not in rules_hit(result)


def test_email_content_line_starting_like_a_closing_gets_the_closing(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "Hi Bob,\nBest option is B.")

    assert result.ok
    assert [fix.rule for fix in result.fixes] == ["signoff.appended"]


def test_rsvp_has_no_closing_rule(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "Running late, will join at 10.", Channel.RSVP)

    assert result.ok
    assert result.fixes == ()


@pytest.mark.parametrize(
    ("line", "expected"),
    [
        ("Best regards,", True),
        ("thanks a lot!", True),
        ("Many thanks, Jane", True),
        ("BR", True),
        ("مع خالص التحية", True),
        ("Thanks for sending the budget numbers yesterday", False),
        ("Thanks for the update", False),
        ("Best option is B.", False),
        ("Best regards, Alex", True),
        ("Thanks, Alex", True),
        ("Best, Alex Example", True),
        ("Bestow the honor", False),
        ("The budget is approved.", False),
        ("", False),
    ],
)
def test_sign_off_like_detection(line: str, expected: bool) -> None:
    assert is_sign_off_like(line) is expected
