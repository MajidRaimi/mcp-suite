from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from office_mcp.features.outlook.compose.gateway import ComposeGateway, ComposeSurface
from office_mcp.features.outlook.compose.guard import SendGuard, SentRecord
from office_mcp.features.outlook.compose.session import run_scoped
from office_mcp.features.outlook.compose.verify import base_subject, sent_recently
from office_mcp.features.outlook.models import Row
from office_mcp.shared.errors import ActionNotVerifiedError, TargetNotFoundError
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.scope import ScopeFactory, ScopeOptions
from office_mcp.shared.waiting import Waiter

UNDO_WINDOW = 30.0
REOPEN_TIMEOUT = 10.0
UNDO_SCOPE = ScopeOptions(foreground=True)
NEWER_MINUTES = 2
UNDONE_DETAIL = (
    "Outlook took the message back before it went out; the reopened draft was discarded and "
    "a copy is in Deleted Items"
)


@dataclass(frozen=True)
class UndoResult:
    subject: str
    detail: str
    focus: FocusOutcome
    notes: tuple[str, ...] = ()

    def to_record(self) -> dict[str, Any]:
        return {
            "ok": True,
            "undone_subject": self.subject,
            "detail": self.detail,
            "focus_restored": self.focus.value,
            "notes": list(self.notes),
        }


class UndoSend:
    def __init__(
        self,
        gateway: ComposeGateway,
        scope: ScopeFactory,
        sends: SendGuard,
        waiter: Waiter,
        outgoing: Callable[[], Sequence[Row]],
        now: Callable[[], datetime],
    ) -> None:
        self._gateway = gateway
        self._scope = scope
        self._sends = sends
        self._waiter = waiter
        self._outgoing = outgoing
        self._now = now

    def undo(self) -> UndoResult:
        record = self._sends.last(UNDO_WINDOW)
        if record is None:
            raise TargetNotFoundError(
                f"no mail was sent by this server in the last {UNDO_WINDOW:.0f} s, so there is "
                "nothing to undo"
            )
        scoped = run_scoped(self._scope, UNDO_SCOPE, lambda _ui: self._take_back(record))
        return UndoResult(record.subject, UNDONE_DETAIL, scoped.focus, scoped.notes)

    def _take_back(self, record: SentRecord) -> None:
        if not self._gateway.undo_send_enabled():
            raise ActionNotVerifiedError(
                "Outlook no longer offers Undo Send, so the mail has most likely gone out; "
                "check Sent Items"
            )
        self._require_newest(record)
        before = self._gateway.surfaces()
        self._gateway.undo_send()
        surface = self._waiter.until(lambda: self._reopened(before), REOPEN_TIMEOUT)
        if surface is None:
            raise ActionNotVerifiedError(
                "Undo Send was pressed but no compose window reopened; check Sent Items and "
                "Drafts before sending again"
            )
        shown = self._gateway.subject(surface)
        if base_subject(shown) != base_subject(record.subject):
            raise ActionNotVerifiedError(
                f"a compose window reopened with the subject {shown[:120]!r} instead of "
                f"{record.subject[:120]!r}; it was left open for you to check"
            )
        self._gateway.discard(surface)
        self._sends.forget(record)

    def _require_newest(self, record: SentRecord) -> None:
        now = self._now()
        wanted = base_subject(record.subject)
        for row in self._outgoing():
            if base_subject(row.summary.title) != wanted and sent_recently(
                row.summary, now, NEWER_MINUTES
            ):
                raise ActionNotVerifiedError(
                    "another mail was sent after this one, and Undo Send would take that one "
                    "back instead; nothing was undone"
                )

    def _reopened(self, before: list[ComposeSurface]) -> ComposeSurface | None:
        return next(iter(self._gateway.stray_surfaces(before)), None)
