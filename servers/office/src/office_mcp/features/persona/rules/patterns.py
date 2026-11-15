from __future__ import annotations

import re
from collections.abc import Iterable, Sequence

from office_mcp.features.persona.rules.common import MAX_HITS_PER_FIELD, TextField, excerpt_around
from office_mcp.shared.drafts import Severity, Violation

RULE = "patterns"
HINT = "Rewrite that part."


def check(targets: Iterable[TextField], patterns: Sequence[re.Pattern[str]]) -> list[Violation]:
    return [
        violation(target, pattern, found)
        for target in targets
        for pattern in patterns
        for found in list(pattern.finditer(target.text))[:MAX_HITS_PER_FIELD]
    ]


def violation(target: TextField, pattern: re.Pattern[str], found: re.Match[str]) -> Violation:
    excerpt = excerpt_around(target.text, found.start(), max(found.end(), found.start() + 1))
    return Violation(
        rule=RULE,
        message=f'text matches the forbidden pattern "{pattern.pattern}" in {target.name}: '
        f'"{excerpt}"',
        severity=Severity.ERROR,
        excerpt=excerpt,
        hint=HINT,
    )
