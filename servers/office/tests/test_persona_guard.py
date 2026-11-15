from __future__ import annotations

from pathlib import Path

import pytest
from fakes.p7_persona import FULL_FRONT_MATTER, guard_for, persona_store, write_persona

from office_mcp.features.persona.guard import PersonaGuard
from office_mcp.shared.drafts import Channel, Draft, TextGuard
from office_mcp.shared.errors import (
    PersonaFormatError,
    PersonaMissingError,
    PersonaViolationError,
    TextEntryError,
    error_text,
)

EM_DASH = "\u2014"
EN_DASH = "\u2013"


def email(body: str, subject: str = "") -> Draft:
    return Draft(
        channel=Channel.EMAIL, body=body, subject=subject, sender_account="alex@example.com"
    )


def test_guard_satisfies_the_text_guard_protocol(tmp_path: Path) -> None:
    guard: TextGuard = guard_for(tmp_path)

    assert guard.review(email("Approved.")).ok


def test_require_clean_returns_the_final_draft(tmp_path: Path) -> None:
    review = guard_for(tmp_path).require_clean(email("Approved."))

    assert review.draft.body == "Approved.\n\nBest regards,"
    assert review.persona_version == guard_for(tmp_path).persona_version


def test_require_clean_raises_with_every_error(tmp_path: Path) -> None:
    with pytest.raises(PersonaViolationError) as caught:
        guard_for(tmp_path).require_clean(email(f"The plan {EM_DASH} then\n\nBest regards,\nAlex"))

    assert [item.rule for item in caught.value.violations] == ["signoff.trailing", "characters"]


def test_error_text_format(tmp_path: Path) -> None:
    body = f"We agreed on the plan {EM_DASH} then the budget.\n\nBest regards,\nAlex"

    with pytest.raises(PersonaViolationError) as caught:
        guard_for(tmp_path).require_clean(email(body))

    assert error_text(caught.value) == (
        "Not sent: the text breaks 2 persona rules. Fix them and call the tool again.\n"
        '1. signoff.trailing: nothing may follow "Best regards," (the signature card is attached '
        'automatically). Found "Alex".\n'
        '2. characters: forbidden character U+2014 EM DASH in body: "...ed on the plan [U+2014] '
        'then the budge...". Use a comma, period or colon.'
    )
    assert EM_DASH not in str(caught.value)


def test_warnings_do_not_block(tmp_path: Path) -> None:
    review = guard_for(tmp_path).require_clean(email(" ".join(["word"] * 60)))

    assert review.ok
    assert review.warnings


def test_verify_written_accepts_the_exact_text(tmp_path: Path) -> None:
    guard = guard_for(tmp_path)
    final = guard.require_clean(email("Approved.", subject="Project Atlas budget")).draft

    guard.verify_written(final, "Approved.  \r\n\r\nBest regards,\n", "Project Atlas budget")


def test_verify_written_catches_a_substituted_dash(tmp_path: Path) -> None:
    guard = guard_for(tmp_path)
    final = guard.require_clean(email("Budget for 2025-2026.")).draft

    with pytest.raises(TextEntryError) as caught:
        guard.verify_written(final, f"Budget for 2025{EN_DASH}2026.\n\nBest regards,", "")

    assert "[U+2013]" in str(caught.value)


def test_verify_written_catches_a_joined_closing_line(tmp_path: Path) -> None:
    guard = guard_for(tmp_path)
    final = guard.require_clean(email("Approved.")).draft

    with pytest.raises(TextEntryError):
        guard.verify_written(final, "Approved. Best regards,", "")


def test_verify_written_rechecks_the_rules(tmp_path: Path) -> None:
    guard = guard_for(tmp_path)
    intended = Draft(channel=Channel.TEAMS, body=f"plan {EM_DASH} then")

    with pytest.raises(PersonaViolationError):
        guard.verify_written(intended, f"plan {EM_DASH} then", "")


def test_missing_persona_blocks_review(tmp_path: Path) -> None:
    guard = PersonaGuard(persona_store(tmp_path / "persona.md", explicit=True))

    with pytest.raises(PersonaMissingError):
        guard.review(email("Approved."))
    with pytest.raises(PersonaMissingError):
        guard.require_clean(email("Approved."))


def test_invalid_persona_blocks_review(tmp_path: Path) -> None:
    path = write_persona(tmp_path / "persona.md", "+++\n[email]\nsign_of = 1\n+++\n")
    guard = PersonaGuard(persona_store(path))

    with pytest.raises(PersonaFormatError, match=r"line 3: email\.sign_of"):
        guard.require_clean(email("Approved."))


def test_default_missing_persona_allows_everything(tmp_path: Path) -> None:
    guard = PersonaGuard(persona_store(tmp_path / "persona.md", explicit=False))

    review = guard.require_clean(email(f"plan {EM_DASH} then"))

    assert review.ok
    assert review.persona_version == "none"


def test_edit_applies_without_restart(tmp_path: Path) -> None:
    path = write_persona(tmp_path / "persona.md", "voice only")
    guard = PersonaGuard(persona_store(path))
    assert guard.review(email(f"plan {EM_DASH} then")).ok

    write_persona(path, FULL_FRONT_MATTER)

    assert not guard.review(email(f"plan {EM_DASH} then")).ok
