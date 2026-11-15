from __future__ import annotations

import re
from collections.abc import Sequence
from datetime import datetime, timedelta

from office_mcp.features.outlook.compose.gateway import ReplyKind
from office_mcp.features.outlook.models import Row, Summary
from office_mcp.shared.errors import ComposeBusyError, ComposeError, TextEntryError
from office_mcp.shared.text import normalize

SUBJECT_PREFIX = re.compile(r"^\s*(re|fw|fwd)\s*:\s*", re.IGNORECASE)
REPLY_PREFIXES = frozenset({"re"})
FORWARD_PREFIXES = frozenset({"fw", "fwd"})
CLOCK_TIME = re.compile(r"^(?P<hour>\d{1,2}):(?P<minute>\d{2})(?:\s*(?P<half>[AP]M))?$", re.I)
SAFE_ALERT_BUTTONS = ("Cancel", "Don't Send", "Continue Editing")
INVISIBLE = "￼​⁠﻿"
TOP_ROWS = 5
RECENT_MINUTES = 10
CLOCK_SKEW = timedelta(minutes=1)
DISPLAY_PREFIX = {ReplyKind.REPLY: "RE: ", ReplyKind.REPLY_ALL: "RE: ", ReplyKind.FORWARD: "FW: "}


def base_subject(subject: str) -> str:
    text = subject
    while True:
        stripped = SUBJECT_PREFIX.sub("", text, count=1)
        if stripped == text:
            return normalize(stripped)
        text = stripped


def first_prefix(subject: str) -> str:
    found = SUBJECT_PREFIX.match(subject)
    return found.group(1).casefold() if found is not None else ""


def display_subject(kind: ReplyKind, original: str) -> str:
    text = original
    while (stripped := SUBJECT_PREFIX.sub("", text, count=1)) != text:
        text = stripped
    return f"{DISPLAY_PREFIX[kind]}{text.strip()}"


def require_subject(kind: ReplyKind, original: str, shown: str) -> None:
    allowed = FORWARD_PREFIXES if kind is ReplyKind.FORWARD else REPLY_PREFIXES
    if first_prefix(shown) not in allowed or base_subject(shown) != base_subject(original):
        raise ComposeError(
            f"the compose subject {shown[:120]!r} does not match "
            f"{display_subject(kind, original)[:120]!r}"
        )


def body_is_free(baseline: str) -> bool:
    lines = baseline.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    return not lines[0].strip(f" \t{INVISIBLE}")


def require_free_body(baseline: str) -> None:
    if not body_is_free(baseline):
        raise ComposeBusyError(
            "the compose body already holds text above the signature (perhaps an unfinished "
            "draft); finish or discard it in Outlook first"
        )


def written_part(full: str, baseline: str) -> str:
    if not baseline:
        return full
    if not full.endswith(baseline):
        raise TextEntryError(
            "the signature or quoted text below the body changed while writing, so the written "
            "text cannot be checked"
        )
    return full[: len(full) - len(baseline)]


def safe_alert_button(buttons: Sequence[str]) -> str | None:
    return next((title for title in SAFE_ALERT_BUTTONS if title in buttons), None)


def title_has_subject(title: str, subject: str) -> bool:
    wanted = normalize(subject)
    shown = normalize(title)
    return bool(wanted) and (shown == wanted or shown.endswith(f" {wanted}"))


def sent_at(summary: Summary, now: datetime) -> datetime | None:
    if summary.date != now.date().isoformat():
        return None
    match = CLOCK_TIME.match(summary.time_label.strip())
    if match is None:
        return now
    hour = int(match["hour"]) % 24
    half = (match["half"] or "").upper()
    if half == "PM" and hour < 12:
        hour += 12
    if half == "AM" and hour == 12:
        hour = 0
    return now.replace(hour=hour, minute=int(match["minute"]), second=0, microsecond=0)


def sent_recently(summary: Summary, now: datetime, minutes: int = RECENT_MINUTES) -> bool:
    when = sent_at(summary, now)
    if when is None:
        return False
    age = now - when
    return -CLOCK_SKEW <= age <= timedelta(minutes=minutes)


def matching_rows(rows: Sequence[Row], subject: str) -> list[Row]:
    return [row for row in rows[:TOP_ROWS] if title_has_subject(row.summary.title, subject)]


def recent_duplicates(rows: Sequence[Row], subject: str, now: datetime) -> list[Row]:
    return [row for row in matching_rows(rows, subject) if sent_recently(row.summary, now)]


def shows_topic(title: str, subject: str) -> bool:
    return title_has_subject(title, subject) or title_has_subject(title, base_subject(subject))


def new_matches(rows: Sequence[Row], subject: str, known: frozenset[str]) -> list[Row]:
    return [
        row
        for row in rows[:TOP_ROWS]
        if shows_topic(row.summary.title, subject) and row.summary.key not in known
    ]
