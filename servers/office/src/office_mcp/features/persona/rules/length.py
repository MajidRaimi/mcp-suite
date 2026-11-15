from __future__ import annotations

from office_mcp.features.persona.models import ChannelRules
from office_mcp.shared.drafts import Severity, Violation

RULE = "length"


def check(body: str, rules: ChannelRules) -> list[Violation]:
    found: list[Violation] = []
    words = len(body.split())
    if rules.max_words is not None and words > rules.max_words:
        found.append(
            warning(f"the text has {words} words; the persona asks for at most {rules.max_words}")
        )
    if rules.max_chars is not None and len(body) > rules.max_chars:
        found.append(
            warning(
                f"the text has {len(body)} characters; the persona asks for at most "
                f"{rules.max_chars}"
            )
        )
    return found


def warning(message: str) -> Violation:
    return Violation(rule=RULE, message=message, severity=Severity.WARNING, hint="Shorten it.")
