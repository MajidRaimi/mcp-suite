from __future__ import annotations

from pathlib import Path

from office_mcp.features.persona.models import CharacterSet, SignOffRule
from office_mcp.features.persona.schema import CHANNEL_KEYS, FORBID_KEYS, LANGUAGE_KEYS
from office_mcp.features.persona.store import parse_persona
from office_mcp.shared.drafts import Channel

EXAMPLE = Path(__file__).resolve().parents[1] / "examples" / "persona.example.md"
HEADINGS = ("# Voice", "# Language", "# Email", "# Teams", "# Do", "# Do not", "# Examples")


def test_example_loads_without_errors() -> None:
    persona = parse_persona(EXAMPLE.read_bytes(), EXAMPLE)

    rules = persona.rules
    assert rules.forbid.character_sets == (CharacterSet.DASHES,)
    assert rules.forbid.emoji is True
    assert rules.channels[Channel.EMAIL].sign_off_rule is SignOffRule.REQUIRED
    assert rules.channels[Channel.EMAIL].applies_to_accounts == ("@example.com",)
    assert rules.channels[Channel.TEAMS].sign_off_rule is SignOffRule.FORBIDDEN
    assert set(rules.channels) == set(Channel)


def test_example_lists_every_key_once() -> None:
    text = EXAMPLE.read_text(encoding="utf-8")
    front = text.split("+++")[1]
    keys = [line.split("=")[0].strip() for line in front.splitlines() if "=" in line]

    for key in (*LANGUAGE_KEYS, *FORBID_KEYS, "version"):
        assert keys.count(key) == 1, key
    assert set(CHANNEL_KEYS) <= set(keys)


def test_example_prose_has_the_expected_headings() -> None:
    lines = parse_persona(EXAMPLE.read_bytes(), EXAMPLE).prose.splitlines()

    assert [line for line in lines if line.startswith("# ")] == list(HEADINGS)


def test_example_contains_no_dash_characters() -> None:
    text = EXAMPLE.read_text(encoding="utf-8")

    assert "\u2014" not in text
    assert "\u2013" not in text
