from __future__ import annotations

import re
from collections.abc import Iterable

from office_mcp.features.persona.rules.common import MAX_HITS_PER_FIELD, TextField, excerpt_around
from office_mcp.shared.drafts import Severity, Violation

RULE = "emoticons"
HINT = "Teams turns these into emojis; remove them."
FACES = (":-)", ":)", ":-(", ":(", ":-D", ":D", ";-)", ";)", ":-P", ":P", ":O", ":|")
SYMBOLS = ("(y)", "<3")
LETTER_FACES = ("B)",)


def alternatives(items: tuple[str, ...]) -> str:
    return "|".join(re.escape(item) for item in items)


EMOTICON = re.compile(
    r"(?:(?<![:;<(/])(?i:"
    + alternatives(FACES)
    + r")|(?<![\w:;<(])(?i:"
    + alternatives(SYMBOLS)
    + r")|(?<!\S)(?:"
    + alternatives(LETTER_FACES)
    + r"))(?=$|[\s.,!?])",
    re.MULTILINE,
)


def check(targets: Iterable[TextField]) -> list[Violation]:
    return [
        violation(target, found)
        for target in targets
        for found in list(EMOTICON.finditer(target.text))[:MAX_HITS_PER_FIELD]
    ]


def violation(target: TextField, found: re.Match[str]) -> Violation:
    excerpt = excerpt_around(target.text, found.start(), found.end())
    return Violation(
        rule=RULE,
        message=f'emoticon "{found.group()}" in {target.name}: "{excerpt}"',
        severity=Severity.ERROR,
        excerpt=excerpt,
        hint=HINT,
    )
