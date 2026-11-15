from __future__ import annotations

from office_mcp.features.teams.writing.context import WriteContext, bounded
from office_mcp.features.teams.writing.targets import (
    message_body,
    require_body,
    require_message_id,
)
from office_mcp.features.teams.writing.verify import collapse, looks_deleted
from office_mcp.shared.errors import ElementNotFoundError, ForbiddenActionError
from office_mcp.shared.node import Node
from office_mcp.shared.results import ActionResult, ItemResult
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import UiScope

ACTION = "delete_message"
DELETE_ITEM = "Delete"
DELETE_ITEMS = ("Delete this message", DELETE_ITEM)
DIALOG_ROLES = frozenset({"AXApplicationDialog", "AXSheet"})
DIALOG_SUBROLES = frozenset({"AXApplicationDialog", "AXDialog"})
CONFIRM_ROLE = "AXButton"
CONFIRM_WAIT = 2.0
VERIFY_WAIT = 6.0
LABEL_CHARS = 80


def is_dialog(node: Node) -> bool:
    return node.role in DIALOG_ROLES or node.subrole in DIALOG_SUBROLES


def confirm_button(root: Node) -> Node | None:
    for dialog in (node for node in root.walk() if is_dialog(node)):
        button = dialog.find(lambda node: node.role == CONFIRM_ROLE and node.label() == DELETE_ITEM)
        if button is not None:
            return button
    return None


def removed(root: Node, message_id: str) -> bool:
    body = message_body(root, message_id)
    return body is None or looks_deleted(body.label())


def short_label(text: str) -> str:
    shown = collapse(text)
    return shown if len(shown) <= LABEL_CHARS else shown[: LABEL_CHARS - 3] + "..."


class DeletionService:
    def __init__(self, context: WriteContext) -> None:
        self._context = context

    def delete(self, chat: str, message_id: str, deadline: Deadline) -> ActionResult:
        wanted = require_message_id(message_id)
        scoped = self._context.run(lambda scope: self._delete(scope, chat, wanted, deadline))
        return ActionResult(
            action=ACTION,
            items=(scoped.value,),
            skipped=(),
            focus=scoped.focus,
            notes=scoped.notes,
        )

    def _delete(self, scope: UiScope, chat: str, message_id: str, deadline: Deadline) -> ItemResult:
        targets = self._context.targets
        row = targets.resolve(chat)
        self._context.recipients.require_thread(row.chat.thread_id)
        opened, message = targets.locate(targets.open(row), message_id)
        if not message.mine:
            raise ForbiddenActionError(
                f"the message {message_id} is from {message.author or 'someone else'}; only the "
                "owner's own messages can be deleted"
            )
        scope.ensure_front()
        body = require_body(targets.now(opened.thread_id), message_id)
        item = self._context.menu_item(
            body,
            DELETE_ITEMS,
            lambda: ElementNotFoundError(f"Teams offers no Delete for the message {message_id}"),
        )
        self._context.gateway.press(item)
        self._confirm()
        verified = bool(
            self._context.waiter.until(
                lambda: removed(targets.now(opened.thread_id), message_id),
                bounded(VERIFY_WAIT, deadline),
            )
        )
        return ItemResult(
            ref=message_id,
            ok=verified,
            label=short_label(message.text),
            location=opened.name,
            verified=verified,
            detail="" if verified else "Teams still shows the message; check the chat",
        )

    def _confirm(self) -> None:
        button = self._context.waiter.until(
            lambda: confirm_button(self._context.gateway.snapshot()), CONFIRM_WAIT
        )
        if button is not None:
            self._context.gateway.press(button)
