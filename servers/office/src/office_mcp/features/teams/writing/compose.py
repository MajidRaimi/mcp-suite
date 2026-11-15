from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.messages.parse import message_pane
from office_mcp.shared.drafts import written_exactly
from office_mcp.shared.errors import (
    ComposeBusyError,
    ComposeError,
    ElementNotFoundError,
    OfficeError,
    TextEntryError,
)
from office_mcp.shared.events import Key, Modifier, Shortcut
from office_mcp.shared.node import Node
from office_mcp.shared.text_input import TextWriter

BOX_ROLE = "AXTextArea"
BOX_PREFIX = "new-message-"
PLACEHOLDER = "Type a message"
TOOLBAR_ROLE = "AXToolbar"
SEND_TOOLBAR = "Actions for new message"
BUTTON_ROLE = "AXButton"
SEND_TITLE = "Send"
SEND_SHORTCUT = Shortcut(Key.RETURN, Modifier.COMMAND)
QUOTE_PREFIX = "Begin quote"
NO_KEYBOARD_TARGET = (
    "the Send button was not found and the compose box does not hold keyboard focus in the "
    "main window, so the send shortcut was not pressed; nothing was sent"
)
BUSY_MESSAGE = (
    "the chat has an unsent draft in its compose box; finish or clear it first. Nothing was sent."
)


@dataclass(frozen=True)
class Commit:
    button: Node | None

    @property
    def path(self) -> str:
        return "Send button" if self.button is not None else SEND_SHORTCUT.describe()


@dataclass(frozen=True)
class QuoteState:
    in_box: str


def compose_boxes(root: Node) -> list[Node]:
    return [
        node for node in root.walk() if node.role == BOX_ROLE and node.dom_id.startswith(BOX_PREFIX)
    ]


def placeholder_of(box: Node) -> str:
    return box.title or box.description or PLACEHOLDER


def send_button(root: Node) -> Node | None:
    toolbar = root.find(lambda node: node.role == TOOLBAR_ROLE and node.label() == SEND_TOOLBAR)
    if toolbar is None:
        return None
    return toolbar.find(
        lambda node: (
            node.role == BUTTON_ROLE and node.label() == SEND_TITLE and node.enabled is not False
        )
    )


def choose_commit(root: Node) -> Commit:
    return Commit(button=send_button(root))


def quote_block(root: Node) -> Node | None:
    pane = message_pane(root)
    inside = {id(node) for node in pane.walk()} if pane is not None else set()
    return root.find(lambda node: id(node) not in inside and node.label().startswith(QUOTE_PREFIX))


def strip_quote(value: str, quote: str) -> str:
    if not quote:
        return value
    if not value.startswith(quote):
        raise TextEntryError("the quote in the compose box changed; nothing was sent")
    return value[len(quote) :]


class ComposeBox:
    def __init__(self, gateway: TeamsGateway, text: TextWriter) -> None:
        self._gateway = gateway
        self._text = text

    def locate(self, root: Node) -> Node:
        boxes = compose_boxes(root)
        if not boxes:
            raise ElementNotFoundError("the chat has no compose box; nothing was sent")
        if len(boxes) > 1:
            raise ComposeError("the chat shows more than one compose box; nothing was sent")
        return boxes[0]

    def read(self, box: Node) -> str:
        return self._text.read(self._gateway.element(box), placeholder_of(box))

    def require_empty(self, box: Node) -> None:
        if self.read(box).strip():
            raise ComposeBusyError(BUSY_MESSAGE)

    def write(self, box: Node, text: str) -> None:
        self._text.paste(self._gateway.element(box), text, placeholder_of(box))

    def discard(
        self, box: Node, quote: str, inserted: str, still_valid: Callable[[], bool]
    ) -> bool:
        try:
            value = self.read(box)
            remainder = strip_quote(value, quote)
            ours = written_exactly(inserted, remainder) or (bool(quote) and not remainder.strip())
            if not ours:
                return False
            return self._text.clear_with_keys(
                self._gateway.element(box), value, still_valid, placeholder_of(box)
            )
        except OfficeError:
            return False

    def is_clear(self, box: Node) -> bool:
        try:
            return not self.read(box).strip()
        except OfficeError:
            return False

    def still_holds(self, box: Node, quote: str, inserted: str) -> bool:
        try:
            return written_exactly(inserted, strip_quote(self.read(box), quote))
        except OfficeError:
            return False

    def require_commit_target(self, box: Node, commit: Commit) -> None:
        if commit.button is not None:
            return
        try:
            self._text.focus(self._gateway.element(box))
        except TextEntryError as error:
            raise ElementNotFoundError(NO_KEYBOARD_TARGET) from error

    def commit(self, commit: Commit) -> None:
        if commit.button is not None:
            self._gateway.press(commit.button)
            return
        self._gateway.press_key(SEND_SHORTCUT)
