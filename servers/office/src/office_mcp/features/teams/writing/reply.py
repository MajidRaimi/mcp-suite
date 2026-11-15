from __future__ import annotations

from office_mcp.features.teams.messages.models import Message
from office_mcp.features.teams.writing.compose import ComposeBox, QuoteState, quote_block
from office_mcp.features.teams.writing.context import WriteContext
from office_mcp.features.teams.writing.targets import require_body
from office_mcp.shared.errors import ComposeError, ElementNotFoundError

REPLY_ITEM = "Reply with quote"
QUOTE_WAIT = 4.0


class Quoter:
    def __init__(self, context: WriteContext, box: ComposeBox) -> None:
        self._context = context
        self._box = box

    def add_quote(self, thread_id: str, message: Message) -> QuoteState:
        body = require_body(self._context.targets.now(thread_id), message.id)
        item = self._context.menu_item(
            body,
            (REPLY_ITEM,),
            lambda: ElementNotFoundError(
                f"Teams offers no {REPLY_ITEM!r} for the message {message.id}; nothing was sent"
            ),
        )
        self._context.gateway.press(item)
        return self._context.waiter.require(
            lambda: self._quoted(thread_id),
            lambda: ComposeError(
                f"Teams did not add a quote of the message {message.id} to the compose box; "
                "nothing was sent"
            ),
            QUOTE_WAIT,
        )

    def _quoted(self, thread_id: str) -> QuoteState | None:
        root = self._context.targets.now(thread_id)
        value = self._box.read(self._box.locate(root))
        if value.strip():
            return QuoteState(in_box=value)
        return QuoteState(in_box="") if quote_block(root) is not None else None
