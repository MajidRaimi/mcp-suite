from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date
from typing import Any

from office_mcp.features.outlook.mail.triage import is_from_owner
from office_mcp.features.outlook.models import Summary
from office_mcp.features.outlook.service import recency
from office_mcp.shared.text import normalize

REPLY_PREFIXES = re.compile(r"^(?:(?:re|fw|fwd|aw|wg)\s*:\s*)+", re.IGNORECASE)
TITLE_SEPARATOR = ", "


@dataclass(frozen=True)
class FollowUp:
    key: str
    subject: str
    sent_date: str
    to_preview: str
    days_waiting: int

    def to_record(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "subject": self.subject,
            "sent_date": self.sent_date,
            "to_preview": self.to_preview,
            "days_waiting": self.days_waiting,
        }


def split_title(title: str) -> tuple[str, str]:
    people, separator, subject = title.partition(TITLE_SEPARATOR)
    return (people, subject) if separator else ("", title)


def subject_key(subject: str) -> str:
    return normalize(REPLY_PREFIXES.sub("", subject.strip()))


def reply_stamps(
    inbox: Sequence[Summary], owner_email: str | None
) -> dict[str, list[tuple[str, int]]]:
    stamps: dict[str, list[tuple[str, int]]] = {}
    for summary in inbox:
        if summary.date is None or is_from_owner(summary, owner_email):
            continue
        key = subject_key(split_title(summary.title)[1])
        stamps.setdefault(key, []).append(recency(summary))
    return stamps


def follow_ups(
    sent: Sequence[Summary],
    inbox: Sequence[Summary],
    owner_email: str | None,
    today: date,
    min_days: int,
    max_days: int,
) -> list[FollowUp]:
    replies = reply_stamps(inbox, owner_email)
    seen: set[str] = set()
    waiting: list[FollowUp] = []
    for summary in sorted(sent, key=recency, reverse=True):
        to_preview, subject = split_title(summary.title)
        key = subject_key(subject)
        if summary.date is None or not key or key in seen:
            continue
        seen.add(key)
        days = (today - date.fromisoformat(summary.date)).days
        answered = any(stamp > recency(summary) for stamp in replies.get(key, []))
        if answered or not min_days <= days <= max_days:
            continue
        waiting.append(FollowUp(summary.key, subject, summary.date, to_preview, days))
    return sorted(waiting, key=lambda item: item.days_waiting, reverse=True)
