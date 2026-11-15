from __future__ import annotations

import unicodedata
from collections.abc import Iterable

from office_mcp.features.persona.models import code_point
from office_mcp.features.persona.rules.common import MAX_HITS_PER_FIELD, TextField, excerpt_around
from office_mcp.shared.drafts import DASH_CHARACTERS, Severity, Violation

RULE = "characters"
DASH_HINT = "Use a comma, period or colon."
OTHER_HINT = "Remove or replace it."


def check(targets: Iterable[TextField], forbidden: frozenset[str]) -> list[Violation]:
    if not forbidden:
        return []
    return [
        violation(target, index)
        for target in targets
        for index in hits(target.text, forbidden)[:MAX_HITS_PER_FIELD]
    ]


def hits(text: str, forbidden: frozenset[str]) -> list[int]:
    return [index for index, char in enumerate(text) if char in forbidden]


def violation(target: TextField, index: int) -> Violation:
    char = target.text[index]
    excerpt = excerpt_around(target.text, index)
    return Violation(
        rule=RULE,
        message=f'forbidden character {describe(char)} in {target.name}: "{excerpt}"',
        severity=Severity.ERROR,
        excerpt=excerpt,
        hint=DASH_HINT if char in DASH_CHARACTERS else OTHER_HINT,
    )


def describe(char: str) -> str:
    name = unicodedata.name(char, "")
    return f"{code_point(char)} {name}" if name else code_point(char)
