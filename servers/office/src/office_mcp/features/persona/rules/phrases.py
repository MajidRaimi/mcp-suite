from __future__ import annotations

from collections.abc import Iterable, Sequence

from office_mcp.features.persona.rules.common import TextField
from office_mcp.shared.drafts import Severity, Violation
from office_mcp.shared.text import normalize

RULE = "phrases"
HINT = "Rewrite that part without the phrase."
TYPOGRAPHIC = str.maketrans(
    {
        "\u2018": "'",
        "\u2019": "'",
        "\u02bc": "'",
        "\uff07": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2010": "-",
        "\u2011": "-",
    }
)


def check(targets: Iterable[TextField], phrases: Sequence[str]) -> list[Violation]:
    folded = [(phrase, fold(phrase)) for phrase in phrases if fold(phrase)]
    return [
        violation(target, phrase)
        for target in targets
        for phrase, needle in folded
        if needle in fold(target.text)
    ]


def fold(text: str) -> str:
    return normalize(text.translate(TYPOGRAPHIC))


def violation(target: TextField, phrase: str) -> Violation:
    return Violation(
        rule=RULE,
        message=f'forbidden phrase "{phrase}" in {target.name}',
        severity=Severity.ERROR,
        excerpt=phrase,
        hint=HINT,
    )
