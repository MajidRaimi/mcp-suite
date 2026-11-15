from __future__ import annotations

import re
from pathlib import Path

import pytest
from fakes.p7_persona import FULL_FRONT_MATTER, persona_store, write_persona

from office_mcp.features.persona.models import (
    DEFAULT_EMOJI_ALLOW,
    CharacterSet,
    Language,
    PersonaState,
    SignOffRule,
)
from office_mcp.features.persona.store import parse_persona
from office_mcp.shared.drafts import Channel
from office_mcp.shared.errors import PersonaFormatError

PATH = Path("/home/alex/.config/office-mcp/persona.md")


def parse(text: str) -> object:
    return parse_persona(text.encode("utf-8"), PATH)


def format_error(text: str) -> str:
    with pytest.raises(PersonaFormatError) as caught:
        parse_persona(text.encode("utf-8"), PATH)
    return str(caught.value)


def test_no_front_matter_is_all_prose_and_no_rules() -> None:
    persona = parse_persona(b"# Voice\nShort and direct.\n", PATH)

    assert persona.prose == "# Voice\nShort and direct."
    assert persona.rules.channels == {}
    assert persona.rules.language is None
    assert persona.rules.forbid.forbidden_characters() == frozenset()


def test_empty_front_matter_is_accepted() -> None:
    persona = parse_persona(b"+++\n+++\nHello", PATH)

    assert persona.prose == "Hello"
    assert persona.rules.version == 1
    assert persona.rules.channels == {}


def test_full_front_matter_parses_every_key() -> None:
    rules = parse_persona(FULL_FRONT_MATTER.encode("utf-8"), PATH).rules

    assert rules.language is not None
    assert rules.language.default is Language.EN
    assert rules.language.mirror == frozenset({Language.AR})
    assert rules.forbid.character_sets == (CharacterSet.DASHES,)
    assert rules.forbid.emoji is True
    assert rules.forbid.emoji_allow == DEFAULT_EMOJI_ALLOW
    assert [item.pattern for item in rules.forbid.patterns] == [r"(?<=\S) - (?=\S)", "--"]
    email = rules.channels[Channel.EMAIL]
    assert email.sign_off_rule is SignOffRule.REQUIRED
    assert email.closing(Language.EN) == "Best regards,"
    assert email.applies_to_accounts == ("@example.com",)
    assert rules.channels[Channel.TEAMS].sign_off_rule is SignOffRule.FORBIDDEN
    assert rules.channels[Channel.RSVP].max_words == 10


def test_unclosed_front_matter_is_an_error() -> None:
    message = format_error("+++\nversion = 1\n# Voice\n")

    assert "line 1" in message
    assert "no closing +++" in message


def test_invalid_toml_reports_the_file_line() -> None:
    message = format_error("+++\nversion = 1\n[forbid]\nemoji = = true\n+++\n")

    assert str(PATH) in message
    assert "line 4" in message
    assert "not valid TOML" in message


def test_unknown_key_names_the_key_path_and_line() -> None:
    message = format_error(
        '+++\n[email]\nsign_off = { default = "Best regards,", fr = "Cordialement," }\n+++\n'
    )

    assert "email.sign_off.fr" in message
    assert "line 3" in message
    assert "not a known key" in message


def test_unknown_top_level_table_is_an_error() -> None:
    message = format_error("+++\n[signature]\nname = 1\n+++\n")

    assert re.search(r"line 2: signature: is not a known key", message)


@pytest.mark.parametrize(
    ("body", "key_path"),
    [
        ('[forbid]\nemoji = "yes"', "forbid.emoji"),
        ("[email]\nmax_words = true", "email.max_words"),
        ("[email]\nmax_words = 0", "email.max_words"),
        ('[language]\ndefault = "fr"', "language.default"),
        ('[language]\nmirror = ["ar", 3]', "language.mirror[1]"),
        ('[forbid]\ncharacters = ["ab"]', "forbid.characters[0]"),
        ('[forbid]\ncharacter_sets = ["quotes"]', "forbid.character_sets[0]"),
        ('[teams]\nsign_off_rule = "sometimes"', "teams.sign_off_rule"),
        ('[email]\nsign_off_rule = "required"', "email.sign_off"),
        ('[email]\napplies_to_accounts = ["example"]', "email.applies_to_accounts[0]"),
        ("version = 2", "version"),
        ('forbid = "dashes"', "forbid"),
    ],
)
def test_wrong_types_name_the_key_path(body: str, key_path: str) -> None:
    message = format_error(f"+++\n{body}\n+++\n")

    assert f"{key_path}:" in message


def test_bad_regex_is_a_format_error() -> None:
    message = format_error('+++\n[forbid]\npatterns = ["ok", "(unclosed"]\n+++\n')

    assert "forbid.patterns[1]" in message
    assert "not a valid regular expression" in message


def test_sign_off_string_becomes_the_default_closing() -> None:
    rules = parse_persona(b'+++\n[teams]\nsign_off = "Thanks"\n+++\n', PATH).rules

    assert rules.channels[Channel.TEAMS].closing(Language.AR) == "Thanks"


def test_emoji_allow_override_replaces_the_default() -> None:
    rules = parse_persona(b"+++\n[forbid]\nemoji = true\nemoji_allow = []\n+++\n", PATH).rules

    assert rules.forbid.emoji_allow == frozenset()


def test_not_utf8_is_invalid(tmp_path: Path) -> None:
    path = tmp_path / "persona.md"
    path.write_bytes(b"\xff\xfe\x00bad")

    status = persona_store(path).status()

    assert status.state is PersonaState.INVALID
    assert status.error is not None and "UTF-8" in status.error


def test_rules_record_is_plain_data(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER))

    record = store.current().rules.to_record()

    assert record["forbid"]["character_sets"] == ["dashes"]
    assert record["forbid"]["emoji_allow"] == ["U+00A9", "U+00AE", "U+2122"]
    assert record["channels"]["email"]["sign_off"]["default"] == "Best regards,"
    assert record["language"] == {"default": "en", "mirror": ["ar"], "enforce": True}
