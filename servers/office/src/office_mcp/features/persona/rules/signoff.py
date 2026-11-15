from __future__ import annotations

import re
from dataclasses import dataclass

from office_mcp.features.persona.models import ChannelRules, Language, SignOffRule
from office_mcp.shared.drafts import Fix, Severity, Violation
from office_mcp.shared.text import normalize

APPENDED = "signoff.appended"
WRONG = "signoff.wrong"
TRAILING = "signoff.trailing"
FORBIDDEN = "signoff.forbidden"
MAX_SIGN_OFF_WORDS = 4
MAX_NAME_WORDS = 2
MIN_LINES_FOR_CLOSING = 2
TRAILING_PUNCTUATION = ",.!، "
SIGN_OFF_STARTS = tuple(
    normalize(phrase)
    for phrase in (
        "regards",
        "best",
        "best regards",
        "kind regards",
        "warm regards",
        "thanks",
        "thank you",
        "many thanks",
        "cheers",
        "sincerely",
        "br",
        "تحياتي",
        "مع التحية",
        "مع خالص التحية",
        "شكرا",
        "وشكرا",
        "مع الشكر",
    )
)
SIGN_OFF_PATTERN = re.compile(
    r"^(?P<phrase>"
    + "|".join(re.escape(phrase) for phrase in sorted(SIGN_OFF_STARTS, key=len, reverse=True))
    + r")(?:\W|$)"
)


@dataclass(frozen=True)
class SignOffOutcome:
    body: str
    violations: tuple[Violation, ...] = ()
    fixes: tuple[Fix, ...] = ()


def is_sign_off_like(line: str) -> bool:
    folded = normalize(line).rstrip(TRAILING_PUNCTUATION).strip()
    if not folded or len(folded.split()) > MAX_SIGN_OFF_WORDS:
        return False
    match = SIGN_OFF_PATTERN.match(folded)
    if match is None:
        return False
    rest = folded[len(match["phrase"]) :].strip(TRAILING_PUNCTUATION)
    return len(rest.split()) <= MAX_NAME_WORDS


def non_empty_lines(body: str) -> list[str]:
    return [line.strip() for line in body.split("\n") if line.strip()]


def apply(body: str, rules: ChannelRules, language: Language) -> SignOffOutcome:
    if rules.sign_off_rule is SignOffRule.FORBIDDEN:
        return SignOffOutcome(body=body, violations=forbidden(body, rules))
    closing = rules.closing(language)
    if closing is None:
        return SignOffOutcome(body=body)
    if rules.nothing_after_sign_off:
        trailing = after_closing(body, closing)
        if trailing is not None:
            return SignOffOutcome(body=body, violations=(trailing,))
    if rules.sign_off_rule is SignOffRule.REQUIRED:
        return required(body, closing)
    return SignOffOutcome(body=body)


def required(body: str, closing: str) -> SignOffOutcome:
    lines = non_empty_lines(body)
    last = lines[-1] if lines else ""
    if last == closing:
        return SignOffOutcome(body=body)
    if lines and is_sign_off_like(last):
        return SignOffOutcome(body=body, violations=(wrong(closing, last),))
    appended = f"{body}\n\n{closing}" if body else closing
    fix = Fix(rule=APPENDED, detail=f'added "{closing}" on its own line at the end')
    return SignOffOutcome(body=appended, fixes=(fix,))


def after_closing(body: str, closing: str) -> Violation | None:
    lines = non_empty_lines(body)
    if closing not in lines:
        return None
    following = lines[lines.index(closing) + 1 :]
    if not following:
        return None
    return Violation(
        rule=TRAILING,
        message=(
            f'nothing may follow "{closing}" (the signature card is attached automatically). '
            f'Found "{following[0]}".'
        ),
        severity=Severity.ERROR,
        excerpt=following[0],
    )


def wrong(closing: str, found: str) -> Violation:
    return Violation(
        rule=WRONG,
        message=f'the text must end with "{closing}" on its own line. Found "{found}".',
        severity=Severity.ERROR,
        excerpt=found,
        hint=f'Replace that line with "{closing}" or remove it.',
    )


def forbidden(body: str, rules: ChannelRules) -> tuple[Violation, ...]:
    lines = non_empty_lines(body)
    if len(lines) < MIN_LINES_FOR_CLOSING:
        return ()
    last = lines[-1]
    if last not in rules.closings() and not is_sign_off_like(last):
        return ()
    return (
        Violation(
            rule=FORBIDDEN,
            message=f'the text must not end with a closing line. Found "{last}".',
            severity=Severity.ERROR,
            excerpt=last,
            hint="Remove the closing.",
        ),
    )
