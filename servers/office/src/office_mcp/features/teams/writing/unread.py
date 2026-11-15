from __future__ import annotations

from office_mcp.features.teams.writing.context import WriteContext, bounded
from office_mcp.features.teams.writing.targets import require_body
from office_mcp.features.teams.writing.verify import newest_from_others
from office_mcp.shared.errors import ActionNotVerifiedError
from office_mcp.shared.results import ActionResult, ItemResult
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import UiScope

ACTION = "mark_unread"
MARK_UNREAD_ITEM = "Mark as unread"
VERIFY_WAIT = 6.0
LOCATION = "chat list"
ONLY_OWN = (
    "Teams offers Mark as unread only on messages from others, and this chat shows none; "
    "nothing was changed"
)


class UnreadService:
    def __init__(self, context: WriteContext) -> None:
        self._context = context

    def mark_unread(self, chat: str, deadline: Deadline) -> ActionResult:
        scoped = self._context.run(lambda scope: self._mark(scope, chat, deadline))
        return ActionResult(
            action=ACTION,
            items=(scoped.value,),
            skipped=(),
            focus=scoped.focus,
            notes=scoped.notes,
        )

    def _mark(self, scope: UiScope, chat: str, deadline: Deadline) -> ItemResult:
        targets = self._context.targets
        opened = targets.open(targets.resolve(chat))
        newest = newest_from_others(opened.messages)
        if newest is None:
            raise ActionNotVerifiedError(ONLY_OWN)
        scope.ensure_front()
        body = require_body(targets.now(opened.thread_id), newest.id)
        item = self._context.menu_item(
            body,
            (MARK_UNREAD_ITEM,),
            lambda: ActionNotVerifiedError(
                f"Teams did not offer Mark as unread on the message {newest.id}; nothing was "
                "changed"
            ),
        )
        self._context.gateway.press(item)
        self._context.gateway.close_popups()
        verified = bool(
            self._context.waiter.until(
                lambda: self._unread(opened.thread_id), bounded(VERIFY_WAIT, deadline)
            )
        )
        return ItemResult(
            ref=opened.thread_id,
            ok=verified,
            label=opened.name,
            location=LOCATION,
            verified=verified,
            detail=f"marked unread from the message {newest.id}"
            if verified
            else "Teams does not show the chat as unread; check the chat list",
        )

    def _unread(self, thread_id: str) -> bool:
        rows = self._context.targets.rows()
        return any(row.chat.thread_id == thread_id and row.chat.unread for row in rows)
