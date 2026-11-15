from __future__ import annotations

from collections.abc import Sequence

from office_mcp.shared.errors import ComposeError, TextEntryError

MAX_LEVELS = 3
MAX_LABEL = 120
SEPARATORS = ("\\", "/", ">", " ")


def validate_sensitivity(path: Sequence[str]) -> tuple[str, ...]:
    labels = tuple(label.strip() for label in path)
    if len(labels) > MAX_LEVELS:
        raise ComposeError(f"sensitivity takes a menu path of at most {MAX_LEVELS} labels")
    for label in labels:
        if not label or len(label) > MAX_LABEL:
            raise ComposeError(f"each sensitivity label takes 1 to {MAX_LABEL} characters")
        if "\n" in label or "\r" in label:
            raise ComposeError("each sensitivity label must be a single line")
    return labels


def fold(text: str) -> str:
    return " ".join(text.split()).casefold()


def shows_label(shown: str, label: str) -> bool:
    text = fold(shown)
    wanted = fold(label)
    if not wanted:
        return False
    return text == wanted or any(text.endswith(f"{mark}{wanted}") for mark in SEPARATORS)


def require_label(shown: Sequence[str], path: tuple[str, ...]) -> None:
    if not path:
        return
    if any(shows_label(text, path[-1]) for text in shown):
        return
    raise TextEntryError(f"the compose does not show the sensitivity label {path[-1]!r}")
