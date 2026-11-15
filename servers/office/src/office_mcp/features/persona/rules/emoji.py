from __future__ import annotations

from collections.abc import Iterable

import regex

from office_mcp.features.persona.models import code_point
from office_mcp.features.persona.rules.common import MAX_HITS_PER_FIELD, TextField, excerpt_around
from office_mcp.shared.drafts import Severity, Violation

RULE = "emoji"
HINT = "Remove emojis and write the words instead."
EMOJI_PART = regex.compile(
    r"[\p{Extended_Pictographic}\p{Regional_Indicator}\p{Emoji_Modifier}\u20e3\ufe0f]"
)
GRAPHEME = regex.compile(r"\X")


def check(targets: Iterable[TextField], allow: frozenset[str]) -> list[Violation]:
    return [
        violation(target, start, end)
        for target in targets
        for start, end in find_emoji(target.text, allow)[:MAX_HITS_PER_FIELD]
    ]


def find_emoji(text: str, allow: frozenset[str]) -> list[tuple[int, int]]:
    found: list[tuple[int, int]] = []
    for cluster in GRAPHEME.finditer(text):
        value = cluster.group()
        if value in allow or EMOJI_PART.search(value) is None:
            continue
        found.append((cluster.start(), cluster.end()))
    return found


def count_emoji(text: str, allow: frozenset[str]) -> int:
    return len(find_emoji(text, allow))


def violation(target: TextField, start: int, end: int) -> Violation:
    points = " ".join(code_point(char) for char in target.text[start:end])
    excerpt = excerpt_around(target.text, start, end)
    return Violation(
        rule=RULE,
        message=f'emoji {points} in {target.name}: "{excerpt}"',
        severity=Severity.ERROR,
        excerpt=excerpt,
        hint=HINT,
    )
