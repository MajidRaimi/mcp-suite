from __future__ import annotations

from pathlib import Path

import pytest
from fakes.p7_persona import guard_for, review, rules_hit

from office_mcp.features.persona.rules.emoji import find_emoji
from office_mcp.shared.drafts import Channel

DASHES = [
    "\u2012",
    "\u2013",
    "\u2014",
    "\u2015",
    "\u2212",
    "\u2e3a",
    "\u2e3b",
    "\ufe31",
    "\ufe32",
    "\ufe58",
    "\ufe63",
    "\uff0d",
]
ARABIC_REPLY = (
    "\u0634\u0643\u0631\u0627 \u0639\u0644\u0649 \u0627\u0644\u062a\u062d\u062f\u064a\u062b"
)
ARABIC_CONTEXT = "\u0647\u0644 \u0627\u0646\u062a\u0647\u064a\u062a \u0645\u0646 \u0645\u064a\u0632\u0627\u0646\u064a\u0629 \u0627\u0644\u0645\u0634\u0631\u0648\u0639"
EMOJI_RULES = """+++
[forbid]
emoji = true
+++
"""
EXAMPLE_PERSONA = Path(__file__).resolve().parents[1] / "examples" / "persona.example.md"
DEFAULT_ALLOW = frozenset({"\u00a9", "\u00ae", "\u2122"})


@pytest.mark.parametrize("dash", DASHES)
def test_every_dash_in_the_preset_is_refused(tmp_path: Path, dash: str) -> None:
    result = review(guard_for(tmp_path), f"The plan {dash} then the budget", Channel.TEAMS)

    assert not result.ok
    assert rules_hit(result) == ["characters"]
    violation = result.violations[0]
    assert violation.hint == "Use a comma, period or colon."
    assert f"U+{ord(dash):04X}" in violation.message
    assert dash not in str(violation.to_record())


def test_dash_in_arabic_text_and_subject_are_both_reported(tmp_path: Path) -> None:
    result = review(
        guard_for(tmp_path),
        f"{ARABIC_REPLY} \u2014 \u0627\u0644\u064a\u0648\u0645",
        Channel.TEAMS,
        context=ARABIC_CONTEXT,
        subject="Budget \u2013 Q3",
    )

    assert [item.message.split(":")[0] for item in result.violations] == [
        "forbidden character U+2013 EN DASH in subject",
        "forbidden character U+2014 EM DASH in body",
    ]


def test_extra_characters_are_refused(tmp_path: Path) -> None:
    guard = guard_for(tmp_path, '+++\n[forbid]\ncharacters = ["\\u2026"]\n+++\n')

    result = review(guard, "Wait\u2026 what")

    assert rules_hit(result) == ["characters"]
    assert "U+2026 HORIZONTAL ELLIPSIS" in result.violations[0].message
    assert result.violations[0].hint == "Remove or replace it."


def test_excerpt_shows_the_surrounding_text(tmp_path: Path) -> None:
    body = "We reviewed everything for the plan \u2014 then we moved on to the budget items"

    violation = review(guard_for(tmp_path), body, Channel.TEAMS).violations[0]

    assert violation.excerpt.startswith("...")
    assert violation.excerpt.endswith("...")
    assert "\u2014" in violation.excerpt
    assert len(violation.excerpt) <= 30 + 6


@pytest.mark.parametrize(
    "text",
    [
        "Great \U0001f44d",
        "Team \U0001f469\u200d\U0001f4bb\u200d",
        "Thanks \U0001f44b\U0001f3fd",
        "Press 1\ufe0f\u20e3 now",
        "From \U0001f1f8\U0001f1e6 today",
        "Logo \u00a9\ufe0f here",
        "\u0634\u0643\u0631\u0627 \u2764\ufe0f",
    ],
)
def test_emoji_forms_are_refused(tmp_path: Path, text: str) -> None:
    result = review(guard_for(tmp_path, EMOJI_RULES), text)

    assert rules_hit(result) == ["emoji"]


def test_zwj_sequence_is_one_hit() -> None:
    family = "\U0001f468\u200d\U0001f469\u200d\U0001f467"

    assert len(find_emoji(f"Hi {family} there", DEFAULT_ALLOW)) == 1


@pytest.mark.parametrize(
    "text", ["Example Corp \u00a9 2026", "Brand\u00ae and Name\u2122", "Plain text"]
)
def test_plain_symbols_are_allowed(tmp_path: Path, text: str) -> None:
    assert review(guard_for(tmp_path, EMOJI_RULES), text).ok


def test_emoji_rule_off_allows_emoji(tmp_path: Path) -> None:
    assert review(guard_for(tmp_path, "+++\n[forbid]\nemoji = false\n+++\n"), "Great \U0001f44d").ok


@pytest.mark.parametrize(
    "emoticon", [":)", ":-)", ":(", ":-(", ":D", ";)", ":P", ":p", ":O", ":|", "(y)", "<3", "B)"]
)
def test_emoticons_refused_on_teams(tmp_path: Path, emoticon: str) -> None:
    result = review(guard_for(tmp_path, EMOJI_RULES), f"Sounds good {emoticon}", Channel.TEAMS)

    assert rules_hit(result) == ["emoticons"]


@pytest.mark.parametrize("text", ["great:)", "ok (Y)", "ok :-D", "fine;-)", "sure :-P"])
def test_attached_and_variant_emoticons_are_refused_on_teams(tmp_path: Path, text: str) -> None:
    result = review(guard_for(tmp_path, EMOJI_RULES), text, Channel.TEAMS)

    assert rules_hit(result) == ["emoticons"]


def test_list_markers_are_not_emoticons(tmp_path: Path) -> None:
    text = "Options:\na) keep it\nb) drop it"

    assert review(guard_for(tmp_path, EMOJI_RULES), text, Channel.TEAMS).ok


def test_emoticons_allowed_in_email(tmp_path: Path) -> None:
    assert review(guard_for(tmp_path, EMOJI_RULES), "Sounds good :)", Channel.EMAIL).ok


def test_emoticon_inside_words_or_links_is_not_matched(tmp_path: Path) -> None:
    text = "See https://example.com/:Docs and the (B) option, 2<3 is true"

    assert review(guard_for(tmp_path, EMOJI_RULES), text, Channel.TEAMS).ok


def test_emoticons_need_the_emoji_rule(tmp_path: Path) -> None:
    assert review(
        guard_for(tmp_path, "+++\n[forbid]\nphrases = []\n+++\n"), "ok :)", Channel.TEAMS
    ).ok


def test_english_phrase_matched_case_and_space_insensitively(tmp_path: Path) -> None:
    result = review(
        guard_for(tmp_path), "Hi Jane,\n\ni HOPE this   email finds you well.", Channel.TEAMS
    )

    assert rules_hit(result) == ["phrases"]
    assert result.violations[0].excerpt == "I hope this email finds you well"


def test_arabic_phrase_with_tashkeel_and_alef_variants_is_matched(tmp_path: Path) -> None:
    variant = "\u0627\u062a\u0645\u0646\u0651\u0649 \u0625\u0646 \u062a\u0643\u0648\u0646\u064e \u0628\u062e\u0640\u064a\u0631\u064d"

    result = review(
        guard_for(tmp_path),
        f"{variant} \u064a\u0627 \u0639\u0645\u0631",
        Channel.TEAMS,
        context=ARABIC_CONTEXT,
    )

    assert rules_hit(result) == ["phrases"]


def test_phrase_written_with_a_curly_apostrophe_is_matched(tmp_path: Path) -> None:
    rules = '+++\n[forbid]\nphrases = ["Don\'t hesitate to reach out"]\n+++\n'

    result = review(guard_for(tmp_path, rules), "Don\u2019t hesitate to reach out", Channel.TEAMS)

    assert rules_hit(result) == ["phrases"]


@pytest.mark.parametrize("dash", ["\u2010", "\u2011", "-"])
def test_example_persona_refuses_spaced_hyphen_variants(tmp_path: Path, dash: str) -> None:
    text = EXAMPLE_PERSONA.read_text(encoding="utf-8")

    result = review(guard_for(tmp_path, text), f"the plan {dash} then", Channel.TEAMS)

    assert rules_hit(result) == ["patterns"]


def test_spaced_hyphen_and_double_hyphen_patterns(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "Plan A - then B -- maybe", Channel.TEAMS)

    assert rules_hit(result) == ["patterns", "patterns"]
    assert "(?<=\\S) - (?=\\S)" in result.violations[0].message


def test_patterns_leave_ordinary_hyphens_alone(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "The follow-up for 2025-2026 is ready", Channel.TEAMS)

    assert result.ok


def test_patterns_apply_to_arabic_text(tmp_path: Path) -> None:
    result = review(
        guard_for(tmp_path),
        f"{ARABIC_REPLY} -- \u063a\u062f\u0627",
        Channel.TEAMS,
        context=ARABIC_CONTEXT,
    )

    assert rules_hit(result) == ["patterns"]


def test_length_warning_does_not_block(tmp_path: Path) -> None:
    body = " ".join(["word"] * 25)

    result = review(guard_for(tmp_path), body, Channel.TEAMS)

    assert result.ok
    assert rules_hit(result) == ["length"]
    assert result.warnings == ("the text has 25 words; the persona asks for at most 20",)


def test_max_chars_warning(tmp_path: Path) -> None:
    guard = guard_for(tmp_path, "+++\n[teams]\nmax_chars = 10\n+++\n")

    result = review(guard, "This is longer than ten", Channel.TEAMS)

    assert result.ok
    assert "23 characters" in result.warnings[0]
