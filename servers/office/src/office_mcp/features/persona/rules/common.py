from __future__ import annotations

from dataclasses import dataclass

EXCERPT_RADIUS = 15
ELLIPSIS = "..."
MAX_HITS_PER_FIELD = 10


@dataclass(frozen=True)
class TextField:
    name: str
    text: str


def fields(body: str, subject: str) -> tuple[TextField, ...]:
    named = (TextField("subject", subject), TextField("body", body))
    return tuple(field for field in named if field.text)


def excerpt_around(text: str, start: int, end: int | None = None) -> str:
    stop = start + 1 if end is None else end
    left = max(0, start - EXCERPT_RADIUS)
    right = min(len(text), stop + EXCERPT_RADIUS)
    window = " ".join(text[left:right].split())
    prefix = ELLIPSIS if left > 0 else ""
    suffix = ELLIPSIS if right < len(text) else ""
    return f"{prefix}{window}{suffix}"
