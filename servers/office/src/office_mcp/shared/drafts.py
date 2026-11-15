from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from typing import Any, Protocol

NO_PERSONA_VERSION = "none"
EXCERPT_CHARS = 60
DASH_CHARACTERS = frozenset(
    {
        "\u2012",
        "\u2013",
        "\u2014",
        "\u2015",
        "\u2212",
        "\u2e3a",
        "\u2e3b",
        "\ufe31",
        "\ufe32",
        "\ufe58",
        "\ufe63",
        "\uff0d",
    }
)


class Channel(StrEnum):
    EMAIL = "email"
    TEAMS = "teams"
    RSVP = "rsvp"


class Severity(StrEnum):
    ERROR = "error"
    WARNING = "warning"


@dataclass(frozen=True)
class Violation:
    rule: str
    message: str
    severity: Severity
    excerpt: str = ""
    hint: str = ""

    def to_record(self) -> dict[str, Any]:
        return {
            "rule": self.rule,
            "message": show_dashes(self.message),
            "severity": self.severity.value,
            "excerpt": show_dashes(self.excerpt),
            "hint": self.hint,
        }


@dataclass(frozen=True)
class Draft:
    channel: Channel
    body: str
    subject: str = ""
    context: str = ""
    context_author: str = ""
    sender_account: str = ""
    recipients: tuple[str, ...] = ()
    allow_language_mismatch: bool = False


@dataclass(frozen=True)
class Fix:
    rule: str
    detail: str

    def to_record(self) -> dict[str, str]:
        return {"rule": self.rule, "detail": self.detail}


@dataclass(frozen=True)
class Review:
    draft: Draft
    violations: tuple[Violation, ...]
    fixes: tuple[Fix, ...]
    persona_version: str

    @property
    def ok(self) -> bool:
        return not any(item.severity is Severity.ERROR for item in self.violations)

    @property
    def warnings(self) -> tuple[str, ...]:
        return tuple(
            show_dashes(item.message)
            for item in self.violations
            if item.severity is Severity.WARNING
        )

    def to_record(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "final_text": self.draft.body,
            "final_subject": self.draft.subject,
            "violations": [item.to_record() for item in self.violations],
            "fixes": [item.to_record() for item in self.fixes],
            "warnings": list(self.warnings),
            "persona_version": self.persona_version,
        }


class TextGuard(Protocol):
    def review(self, draft: Draft) -> Review: ...

    def require_clean(self, draft: Draft) -> Review: ...

    def verify_written(self, intended: Draft, body_read: str, subject_read: str) -> None: ...


class NoTextGuard:
    persona_version = NO_PERSONA_VERSION

    def review(self, draft: Draft) -> Review:
        return Review(draft=draft, violations=(), fixes=(), persona_version=self.persona_version)

    def require_clean(self, draft: Draft) -> Review:
        return self.review(draft)

    def verify_written(self, intended: Draft, body_read: str, subject_read: str) -> None:
        require_written(intended, body_read, subject_read)


def canonical_text(text: str) -> str:
    unified = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip(" \t") for line in unified.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def written_exactly(intended: str, actual: str) -> bool:
    return canonical_text(intended) == canonical_text(actual)


def require_written(intended: Draft, body_read: str, subject_read: str) -> None:
    from office_mcp.shared.errors import TextEntryError

    if intended.subject and not written_exactly(intended.subject, subject_read):
        raise TextEntryError(
            "the subject on screen differs from the checked text: "
            + first_difference(intended.subject, subject_read)
        )
    if not written_exactly(intended.body, body_read):
        raise TextEntryError(
            "the text on screen differs from the checked text: "
            + first_difference(intended.body, body_read)
        )


def first_difference(intended: str, actual: str) -> str:
    expected_lines = canonical_text(intended).split("\n")
    actual_lines = canonical_text(actual).split("\n")
    for number in range(max(len(expected_lines), len(actual_lines))):
        expected = expected_lines[number] if number < len(expected_lines) else ""
        found = actual_lines[number] if number < len(actual_lines) else ""
        if expected != found:
            return f"line {number + 1}: expected {excerpt(expected)}, found {excerpt(found)}"
    return "the texts differ"


def excerpt(text: str, limit: int = EXCERPT_CHARS) -> str:
    shown = text if len(text) <= limit else text[: limit - 3] + "..."
    return f'"{show_dashes(shown)}"'


def show_dashes(text: str) -> str:
    return "".join(f"[U+{ord(char):04X}]" if char in DASH_CHARACTERS else char for char in text)
