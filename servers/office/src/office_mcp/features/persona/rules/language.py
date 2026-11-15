from __future__ import annotations

import re

from office_mcp.features.persona.models import Language, LanguageRules
from office_mcp.shared.drafts import Draft, Severity, Violation
from office_mcp.shared.text import normalize

MISMATCH = "language.mismatch"
NO_CONTEXT = "language.no_context"
ARABIC_THRESHOLD = 0.5
MIN_CONTEXT_LETTERS = 10
CUT_PREFIXES = tuple(
    normalize(prefix)
    for prefix in (
        "disclaimer",
        "confidentiality",
        "this email and any",
        "this message and any",
        "إخلاء",
        "تنبيه",
        "هذه الرسالة",
    )
)
RULE_LINE = re.compile(r"^\s*(?:_{5,}|-{5,})\s*$")
NOISE = (
    re.compile(r"(?:https?://|www\.)\S+", re.IGNORECASE),
    re.compile(r"[\w.%+'-]+@[\w-]+(?:\.[\w-]+)+"),
    re.compile(r"`[^`]*`"),
    re.compile(r"@\w+"),
    re.compile(r"\d+"),
)
ARABIC_RANGES = (
    (0x0600, 0x06FF),
    (0x0750, 0x077F),
    (0x08A0, 0x08FF),
    (0xFB50, 0xFDFF),
    (0xFE70, 0xFEFF),
)
LANGUAGE_NAMES = {Language.EN: "English", Language.AR: "Arabic"}


def language_sample(text: str) -> str:
    kept: list[str] = []
    for line in text.split("\n"):
        if is_cut_line(line):
            break
        kept.append(line)
    sample = "\n".join(kept)
    for pattern in NOISE:
        sample = pattern.sub(" ", sample)
    return sample


def is_cut_line(line: str) -> bool:
    return RULE_LINE.match(line) is not None or normalize(line).startswith(CUT_PREFIXES)


def is_arabic_letter(char: str) -> bool:
    point = ord(char)
    return char.isalpha() and any(low <= point <= high for low, high in ARABIC_RANGES)


def letter_count(text: str) -> int:
    return sum(1 for char in text if char.isalpha())


def arabic_ratio(text: str) -> float:
    letters = letter_count(text)
    if not letters:
        return 0.0
    return sum(1 for char in text if is_arabic_letter(char)) / letters


def language_of(text: str) -> Language:
    return Language.AR if arabic_ratio(text) >= ARABIC_THRESHOLD else Language.EN


def expected_language(context_language: Language, rules: LanguageRules) -> Language:
    return context_language if context_language in rules.mirror else rules.default


def check(draft: Draft, body: str, rules: LanguageRules) -> list[Violation]:
    if draft.allow_language_mismatch:
        return []
    body_sample = language_sample(body)
    if not letter_count(body_sample):
        return []
    written = language_of(body_sample)
    context_sample = language_sample(draft.context)
    if letter_count(context_sample) < MIN_CONTEXT_LETTERS:
        return [] if written is rules.default else [no_context(written, rules.default)]
    context_language = language_of(context_sample)
    wanted = expected_language(context_language, rules)
    if written is wanted:
        return []
    return [mismatch(written, wanted, context_language, draft.context_author, rules.enforce)]


def mismatch(
    written: Language, wanted: Language, context: Language, author: str, enforce: bool
) -> Violation:
    who = author.strip() or "the other side"
    reason = (
        f"{who} wrote in {LANGUAGE_NAMES[context]}"
        if wanted is context
        else f"{LANGUAGE_NAMES[wanted]} is the default language"
    )
    return Violation(
        rule=MISMATCH,
        message=f"the text is in {LANGUAGE_NAMES[written]} but {reason}",
        severity=Severity.ERROR if enforce else Severity.WARNING,
        hint=(
            f"Write it in {LANGUAGE_NAMES[wanted]}, or pass allow_language_mismatch=true when "
            "the owner asked for this language."
        ),
    )


def no_context(written: Language, default: Language) -> Violation:
    return Violation(
        rule=NO_CONTEXT,
        message=(
            f"the text is in {LANGUAGE_NAMES[written]} and the default language is "
            f"{LANGUAGE_NAMES[default]}; make sure the owner asked for it"
        ),
        severity=Severity.WARNING,
    )
