from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any

from office_mcp.features.outlook.models import Summary
from office_mcp.features.outlook.service import recency
from office_mcp.features.outlook.window import FocusedTab
from office_mcp.shared.text import normalize

AUTOMATED_SENDERS = (
    "no-reply",
    "noreply",
    "no_reply",
    "donotreply",
    "do-not-reply",
    "do not reply",
    "notifications",
    "notification",
    "mailer-daemon",
    "postmaster",
)
PINNED_GROUP = "pinned"
DRAFT_FLAG = "draft"
UNREAD_REASON = "unread"
WAITING_REASON = "read, no reply yet"
OTHER_NOTE = "in the Other tab"


@dataclass(frozen=True)
class NeedsReply:
    summary: Summary
    reason: str

    def to_record(self) -> dict[str, Any]:
        return {**self.summary.to_record(), "reason": self.reason}


def sender_part(summary: Summary) -> str:
    return summary.title.split(",", 1)[0]


def is_automated(summary: Summary) -> bool:
    sender = normalize(sender_part(summary))
    return any(marker in sender for marker in AUTOMATED_SENDERS)


def is_from_owner(summary: Summary, owner_email: str | None) -> bool:
    if not owner_email:
        return False
    sender = normalize(sender_part(summary))
    address = owner_email.casefold()
    return address in sender or sender == address.split("@", 1)[0]


def in_window(summary: Summary, today: date, days: int) -> bool:
    since = (today - timedelta(days=days)).isoformat()
    return summary.date is not None and summary.date >= since


def in_other_tab(summary: Summary) -> bool:
    return summary.view == FocusedTab.OTHER.value


def qualifies(summary: Summary, owner_email: str | None, today: date, days: int) -> bool:
    if not in_window(summary, today, days):
        return False
    if summary.replied or summary.forwarded or summary.meeting or DRAFT_FLAG in summary.flags:
        return False
    if is_automated(summary) or is_from_owner(summary, owner_email):
        return False
    if in_other_tab(summary) or summary.group.casefold() == PINNED_GROUP:
        return summary.unread
    return True


def reason_for(summary: Summary) -> str:
    base = UNREAD_REASON if summary.unread else WAITING_REASON
    return f"{base}, {OTHER_NOTE}" if in_other_tab(summary) else base


def mail_needs_reply(
    rows: Sequence[Summary], owner_email: str | None, today: date, days: int
) -> list[NeedsReply]:
    kept = [summary for summary in rows if qualifies(summary, owner_email, today, days)]
    newest = sorted(kept, key=recency, reverse=True)
    ordered = sorted(newest, key=lambda summary: not summary.unread)
    return [NeedsReply(summary, reason_for(summary)) for summary in ordered]
