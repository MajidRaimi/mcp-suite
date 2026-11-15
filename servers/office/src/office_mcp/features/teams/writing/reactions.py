from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from typing import Any

from office_mcp.features.teams.writing.context import WriteContext, bounded
from office_mcp.features.teams.writing.targets import require_body, require_message_id
from office_mcp.shared.errors import ActionNotVerifiedError, ElementNotFoundError
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.node import Node
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import UiScope
from office_mcp.shared.text import normalize

TOGGLE_ROLES = frozenset({"AXCheckBox", "AXToggleButton"})
TOGGLE_SUBROLE = "AXToggleButton"
POPOVER_SUFFIX = "-popover-surface"
ON = "1"
OFF = "0"
POPOVER_WAIT = 3.0
VERIFY_WAIT = 5.0


class Reaction(StrEnum):
    LIKE = "like"
    HEART = "heart"
    LAUGH = "laugh"
    SURPRISED = "surprised"

    @property
    def button_title(self) -> str:
        return self.value.capitalize()


@dataclass(frozen=True)
class ReactionChange:
    message_id: str
    reaction: Reaction
    state: bool
    changed: bool
    verified: bool
    chat: str
    thread_id: str


@dataclass(frozen=True)
class ReactionResult:
    change: ReactionChange
    focus: FocusOutcome
    notes: tuple[str, ...] = ()

    def to_record(self) -> dict[str, Any]:
        change = self.change
        return {
            "ok": change.verified,
            "message_id": change.message_id,
            "reaction": change.reaction.value,
            "state": "on" if change.state else "off",
            "changed": change.changed,
            "verified": change.verified,
            "chat": change.chat,
            "thread_id": change.thread_id,
            "focus_restored": self.focus.value,
            "notes": list(self.notes),
        }


def is_toggle(node: Node, reaction: Reaction) -> bool:
    if node.role not in TOGGLE_ROLES and node.subrole != TOGGLE_SUBROLE:
        return False
    return normalize(node.label()) == normalize(reaction.button_title)


def find_toggle(scope: Node, reaction: Reaction) -> Node | None:
    return scope.find(lambda node: is_toggle(node, reaction))


def popover_toggle(root: Node, message_id: str, reaction: Reaction) -> Node | None:
    surface_id = message_id + POPOVER_SUFFIX
    surface = root.find(lambda node: node.dom_id == surface_id)
    return find_toggle(surface, reaction) if surface is not None else None


def toggle_state(toggle: Node) -> bool | None:
    if toggle.value == ON:
        return True
    if toggle.value == OFF:
        return False
    return None


class ReactionService:
    def __init__(self, context: WriteContext) -> None:
        self._context = context

    def react(
        self, chat: str, message_id: str, reaction: Reaction, remove: bool, deadline: Deadline
    ) -> ReactionResult:
        wanted_id = require_message_id(message_id)
        scoped = self._context.run(
            lambda scope: self._react(scope, chat, wanted_id, reaction, not remove, deadline)
        )
        return ReactionResult(change=scoped.value, focus=scoped.focus, notes=scoped.notes)

    def _react(
        self,
        scope: UiScope,
        chat: str,
        message_id: str,
        reaction: Reaction,
        wanted: bool,
        deadline: Deadline,
    ) -> ReactionChange:
        targets = self._context.targets
        row = targets.resolve(chat)
        self._context.recipients.require_thread(row.chat.thread_id)
        opened, _ = targets.locate(targets.open(row), message_id)
        scope.ensure_front()
        toggle = self._toggle(opened.thread_id, message_id, reaction)
        before = require_state(toggle, reaction)

        def change(state: bool, changed: bool, verified: bool) -> ReactionChange:
            return ReactionChange(
                message_id, reaction, state, changed, verified, opened.name, opened.thread_id
            )

        if before is wanted:
            return change(before, changed=False, verified=True)
        self._context.gateway.press(toggle)
        flipped = self._context.waiter.until(
            lambda: self._state(opened.thread_id, message_id, reaction) is wanted,
            bounded(VERIFY_WAIT, deadline),
        )
        return change(wanted if flipped else before, changed=True, verified=bool(flipped))

    def _toggle(self, thread_id: str, message_id: str, reaction: Reaction) -> Node:
        gateway = self._context.gateway
        body = require_body(self._context.targets.now(thread_id), message_id)
        gateway.focus(body)
        toggle = self._context.waiter.until(
            lambda: popover_toggle(self._context.targets.now(thread_id), message_id, reaction),
            POPOVER_WAIT,
        )
        if toggle is not None:
            return toggle
        toggle = find_toggle(gateway.show_menu(body), reaction)
        if toggle is None:
            gateway.close_popups()
            raise ElementNotFoundError(
                f"Teams shows no {reaction.button_title} reaction for the message {message_id}"
            )
        return toggle

    def _state(self, thread_id: str, message_id: str, reaction: Reaction) -> bool | None:
        root = self._context.targets.now(thread_id)
        toggle = popover_toggle(root, message_id, reaction)
        if toggle is not None:
            return toggle_state(toggle)
        gateway = self._context.gateway
        toggle = find_toggle(gateway.show_menu(require_body(root, message_id)), reaction)
        gateway.close_popups()
        return toggle_state(toggle) if toggle is not None else None


def require_state(toggle: Node, reaction: Reaction) -> bool:
    state = toggle_state(toggle)
    if state is None:
        raise ActionNotVerifiedError(
            f"the state of the {reaction.button_title} reaction could not be read; "
            "nothing was changed"
        )
    return state
