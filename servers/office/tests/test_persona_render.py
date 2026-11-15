from __future__ import annotations

from pathlib import Path

import pytest
from fakes.p7_persona import FULL_FRONT_MATTER, persona_store, write_persona

from office_mcp.features.persona.render import (
    TRUNCATION_MARK,
    instructions_block,
    rules_in_words,
    send_rules_text,
)
from office_mcp.shared.drafts import Channel
from office_mcp.shared.owner import Owner

OWNER = Owner(name="Alex Example")


def test_loaded_block_wording(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER))

    block = instructions_block(store, OWNER)

    assert block.startswith(
        "Everything you write through this server is sent as Alex Example at once, with no "
        "draft step. Write in Alex Example's voice as described below. These hard rules are "
        "checked before anything is sent, and a send that breaks one is rejected with the "
        "list of problems: Email: "
    )
    assert " Teams: " in block
    assert (
        "Call office_persona for the current persona and office_check_text to test a draft "
        "before sending. Persona: # Voice\nShort and direct."
    ) in block
    assert '"Best regards,"' in block
    assert "\u2014" not in block


def test_block_is_capped_and_prose_truncated(tmp_path: Path) -> None:
    prose = "Short and direct. " * 400
    store = persona_store(write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER + prose))

    block = instructions_block(store, OWNER)

    assert len(block) <= 4000
    assert block.endswith(TRUNCATION_MARK)
    assert "Persona: # Voice" in block


def test_small_cap_still_holds(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER))

    assert len(instructions_block(store, OWNER, max_chars=200)) == 200


def test_none_state_block(tmp_path: Path) -> None:
    store = persona_store(tmp_path / "persona.md", explicit=False)

    assert (
        instructions_block(store, OWNER)
        == "No persona is configured, so only the general rules apply."
    )


@pytest.mark.parametrize(
    ("text", "state"),
    [(None, "missing"), ("+++\nversion = 3\n+++\n", "invalid")],
)
def test_blocked_state_block(tmp_path: Path, text: str | None, state: str) -> None:
    path = tmp_path / "persona.md"
    if text is not None:
        write_persona(path, text)

    block = instructions_block(persona_store(path, explicit=True), OWNER)

    assert block == f"The persona file is {state}; sends are blocked until it is fixed."


def test_default_owner_wording(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", "Plain voice."))

    block = instructions_block(store, Owner())

    assert "sent as the user at once" in block
    assert "Write in the user's voice" in block
    assert "Email: no hard rules are configured. Teams: no hard rules are configured." in block


def test_rules_in_words_per_channel(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER))

    email = rules_in_words(store, Channel.EMAIL)
    teams = rules_in_words(store, Channel.TEAMS)
    rsvp = rules_in_words(store, Channel.RSVP)

    assert "no em dash, en dash, minus sign or any other dash character" in email
    assert 'end with "Best regards," (' in email
    assert "when sending from @example.com" in email
    assert "nothing after" in email
    assert "write in English, and in Arabic when the other side writes in it" in email
    assert "aim for at most 50 words" in email
    assert "no emoticons" not in email
    assert "no closing line" in teams
    assert "no emojis and no emoticons" in teams
    assert "Best regards" not in rsvp
    assert "aim for at most 10 words" in rsvp


def test_rules_in_words_when_blocked(tmp_path: Path) -> None:
    store = persona_store(tmp_path / "persona.md", explicit=True)

    assert (
        rules_in_words(store, Channel.EMAIL) == "the persona file is missing, so sends are blocked"
    )


def test_send_rules_text_ends_with_the_check_hint(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER))

    assert send_rules_text(store, Channel.TEAMS).endswith(
        "Call office_check_text first when unsure."
    )
