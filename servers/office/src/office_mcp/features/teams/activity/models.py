from __future__ import annotations

from dataclasses import asdict, dataclass
from enum import StrEnum
from typing import Any


class ActivityKind(StrEnum):
    MENTION = "mention"
    REPLY = "reply"
    REACTION = "reaction"
    INVITE = "invite"
    MEETING = "meeting"
    MISSED_CALL = "missed_call"
    OTHER = "other"


class ActivityFilter(StrEnum):
    ALL = "all"
    UNREAD = "unread"
    MENTIONS = "mentions"


@dataclass(frozen=True)
class ActivityItem:
    kind: ActivityKind
    person: str
    title: str
    preview: str
    time_label: str
    location: str

    def to_record(self) -> dict[str, Any]:
        return {**asdict(self), "kind": self.kind.value}
