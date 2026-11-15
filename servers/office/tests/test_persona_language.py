from __future__ import annotations

from pathlib import Path

from fakes.p7_persona import guard_for, review, rules_hit

from office_mcp.features.persona.rules.language import arabic_ratio, language_sample
from office_mcp.shared.drafts import Channel, Severity

ARABIC_CONTEXT = "هل انتهيت من ميزانية المشروع؟"
ARABIC_REPLY = "نعم، أرسلتها اليوم صباحا"
ENGLISH_CONTEXT = "Did you finish the Project Atlas budget yet?"
DISCLAIMER = (
    "\n\nDisclaimer: This email and any attachments are confidential and intended solely for "
    "the addressee. If you received it in error, please notify the sender and delete it."
)
LENIENT = """+++
[language]
default = "en"
mirror = ["ar"]
enforce = false
+++
"""


def test_english_reply_to_arabic_is_refused_and_names_the_author(tmp_path: Path) -> None:
    result = review(
        guard_for(tmp_path),
        "Yes, I sent it this morning.",
        Channel.TEAMS,
        context=ARABIC_CONTEXT,
        context_author="Omar Haddad",
    )

    assert rules_hit(result) == ["language.mismatch"]
    violation = result.violations[0]
    assert violation.severity is Severity.ERROR
    assert "Omar Haddad wrote in Arabic" in violation.message
    assert "allow_language_mismatch=true" in violation.hint


def test_arabic_reply_to_arabic_is_fine(tmp_path: Path) -> None:
    assert review(guard_for(tmp_path), ARABIC_REPLY, Channel.TEAMS, context=ARABIC_CONTEXT).ok


def test_arabic_reply_to_english_is_refused(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), ARABIC_REPLY, Channel.TEAMS, context=ENGLISH_CONTEXT)

    assert rules_hit(result) == ["language.mismatch"]
    assert (
        result.violations[0].message == "the text is in Arabic but the other side wrote in English"
    )


def test_arabic_context_without_mirror_requires_the_default(tmp_path: Path) -> None:
    rules = '+++\n[language]\ndefault = "en"\n+++\n'

    result = review(guard_for(tmp_path, rules), ARABIC_REPLY, Channel.TEAMS, context=ARABIC_CONTEXT)

    assert rules_hit(result) == ["language.mismatch"]
    assert "English is the default language" in result.violations[0].message


def test_english_disclaimer_under_arabic_text_is_ignored(tmp_path: Path) -> None:
    context = ARABIC_CONTEXT + DISCLAIMER

    assert review(guard_for(tmp_path), ARABIC_REPLY, Channel.TEAMS, context=context).ok
    assert arabic_ratio(language_sample(context)) == 1.0


def test_arabic_disclaimer_and_separator_lines_cut_the_sample() -> None:
    text = "Short note\n_____\n" + ARABIC_CONTEXT
    arabic_cut = "Short note\nإخلاء مسؤولية: " + ARABIC_CONTEXT

    assert language_sample(text).strip() == "Short note"
    assert language_sample(arabic_cut).strip() == "Short note"


def test_urls_emails_and_names_inside_an_arabic_reply(tmp_path: Path) -> None:
    body = (
        f"{ARABIC_REPLY} الرابط https://contoso.com/projects/atlas/budget-2026 "
        "والبريد jane@example.com و@jane_doe مع Jane Doe"
    )

    result = review(guard_for(tmp_path), body, Channel.TEAMS, context=ARABIC_CONTEXT)

    assert result.ok


def test_override_flag_skips_the_rule(tmp_path: Path) -> None:
    result = review(
        guard_for(tmp_path),
        "Yes, sent it.",
        Channel.TEAMS,
        context=ARABIC_CONTEXT,
        allow_language_mismatch=True,
    )

    assert result.ok
    assert result.violations == ()


def test_no_context_warns_when_not_in_the_default(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), ARABIC_REPLY, Channel.TEAMS)

    assert result.ok
    assert rules_hit(result) == ["language.no_context"]
    assert result.violations[0].severity is Severity.WARNING


def test_no_context_in_the_default_is_silent(tmp_path: Path) -> None:
    assert review(guard_for(tmp_path), "Yes, sent it.", Channel.TEAMS).violations == ()


def test_short_context_counts_as_no_context(tmp_path: Path) -> None:
    result = review(guard_for(tmp_path), "Yes, sent it.", Channel.TEAMS, context="تمام 123")

    assert result.violations == ()


def test_not_enforced_mismatch_is_a_warning(tmp_path: Path) -> None:
    result = review(
        guard_for(tmp_path, LENIENT), "Yes, sent it.", Channel.TEAMS, context=ARABIC_CONTEXT
    )

    assert result.ok
    assert result.violations[0].rule == "language.mismatch"
    assert result.violations[0].severity is Severity.WARNING


def test_text_without_letters_is_not_judged(tmp_path: Path) -> None:
    assert review(guard_for(tmp_path), "10:30", Channel.TEAMS, context=ARABIC_CONTEXT).ok


def test_arabic_ratio_counts_letters_only() -> None:
    assert arabic_ratio("") == 0.0
    assert arabic_ratio("123 !!") == 0.0
    assert arabic_ratio("مرحبا ab") == 5 / 7
