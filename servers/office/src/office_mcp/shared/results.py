from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any

from office_mcp.shared.drafts import Channel, Review
from office_mcp.shared.focus import FocusOutcome

UNCONFIRMED_NOTE = "it may have been sent; check Sent Items or the chat before retrying"


@dataclass(frozen=True)
class ItemResult:
    ref: str
    ok: bool
    label: str
    location: str
    verified: bool
    detail: str = ""

    def to_record(self) -> dict[str, Any]:
        return {
            "ref": self.ref,
            "ok": self.ok,
            "label": self.label,
            "location": self.location,
            "verified": self.verified,
            "detail": self.detail,
        }


@dataclass(frozen=True)
class ActionResult:
    action: str
    items: tuple[ItemResult, ...]
    skipped: tuple[str, ...]
    focus: FocusOutcome
    notes: tuple[str, ...] = ()
    extra: Mapping[str, str] = field(default_factory=dict)

    def to_record(self) -> dict[str, Any]:
        return {
            "action": self.action,
            "succeeded": sum(1 for item in self.items if item.ok),
            "failed": sum(1 for item in self.items if not item.ok),
            "skipped": list(self.skipped),
            "results": [item.to_record() for item in self.items],
            "focus_restored": self.focus.value,
            "notes": list(self.notes),
            **self.extra,
        }


class SendStatus(StrEnum):
    SENT = "sent"
    SUBMITTED = "submitted"
    UNCONFIRMED = "unconfirmed"
    DRY_RUN = "dry_run"


OK_STATUSES = frozenset({SendStatus.SENT, SendStatus.SUBMITTED, SendStatus.DRY_RUN})


@dataclass(frozen=True)
class SendResult:
    status: SendStatus
    channel: Channel
    targets: tuple[str, ...]
    subject: str
    message_ref: str
    verified: bool
    review: Review
    focus: FocusOutcome
    notes: tuple[str, ...] = ()
    extra: Mapping[str, Any] = field(default_factory=dict)

    def to_record(self) -> dict[str, Any]:
        return {
            "ok": self.status in OK_STATUSES,
            "status": self.status.value,
            "verified": self.verified,
            "channel": self.channel.value,
            "targets": list(self.targets),
            "subject": self.subject,
            "message_ref": self.message_ref,
            "text_sent": self.review.draft.body,
            "persona": {
                "version": self.review.persona_version,
                "fixes": [fix.to_record() for fix in self.review.fixes],
                "warnings": list(self.review.warnings),
            },
            "focus_restored": self.focus.value,
            "notes": list(self._notes()),
            **self.extra,
        }

    def _notes(self) -> tuple[str, ...]:
        if self.status is SendStatus.UNCONFIRMED and UNCONFIRMED_NOTE not in self.notes:
            return (*self.notes, UNCONFIRMED_NOTE)
        return self.notes
