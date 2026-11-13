from __future__ import annotations

from datetime import tzinfo

from office_mcp.features.teams.messages.models import Message
from office_mcp.shared.clock import message_time
from office_mcp.shared.node import Node

PANE_ID = "chat-pane-list"
BODY_PREFIX = "message-body-"
AUTHOR_PREFIX = "author-"
TIME_PREFIX = "timestamp-"
CONTENT_PREFIX = "content-"
MY_BODY_CLASS = "fui-ChatMyMessage__body"
HEADING_ROLE = "AXHeading"
MESSAGE_HEADING_LEVEL = "4"
BY = " by "
WINDOW_SEPARATOR = " | "


def message_pane(root: Node) -> Node | None:
    return root.find(lambda node: node.dom_id == PANE_ID)


def open_chat_name(root: Node) -> str:
    _, separator, name = root.title.partition(WINDOW_SEPARATOR)
    return name.strip() if separator else ""


def first_content(root: Node) -> Node | None:
    pane = message_pane(root)
    return pane.find(lambda node: node.dom_id.startswith(CONTENT_PREFIX)) if pane else None


def parse_messages(root: Node, zone: tzinfo) -> list[Message]:
    pane = message_pane(root)
    if pane is None:
        return []
    by_id = {node.dom_id: node for node in pane.walk() if node.dom_id}
    messages: list[Message] = []
    heading_author = ""
    for node in pane.walk():
        if node.role == HEADING_ROLE and node.value == MESSAGE_HEADING_LEVEL and BY in node.title:
            heading_author = node.title.rsplit(BY, 1)[1].strip()
        elif node.dom_id.startswith(BODY_PREFIX):
            message_id = node.dom_id.removeprefix(BODY_PREFIX)
            previous = messages[-1].author if messages else ""
            messages.append(
                build_message(message_id, node, by_id, heading_author or previous, zone)
            )
            heading_author = ""
    return messages


def build_message(
    message_id: str, body: Node, by_id: dict[str, Node], fallback: str, zone: tzinfo
) -> Message:
    author_node = by_id.get(AUTHOR_PREFIX + message_id)
    stamp_node = by_id.get(TIME_PREFIX + message_id)
    content_node = by_id.get(CONTENT_PREFIX + message_id)
    author = author_node.text() if author_node else fallback
    stamp = stamp_node.label().rstrip(".") if stamp_node else ""
    sent = message_time(message_id, zone)
    return Message(
        id=message_id,
        author=author,
        mine=body.has_class(MY_BODY_CLASS),
        sent_at=sent.isoformat(timespec="seconds") if sent else stamp,
        text=message_text(body, content_node, author, stamp),
    )


def message_text(body: Node, content: Node | None, author: str, stamp: str) -> str:
    if content is not None:
        return content.label()
    text = body.label()
    suffixes = (f"{author} {stamp}.", f"{author} {stamp}") if stamp else (author,)
    for suffix in suffixes:
        text = text.removesuffix(suffix).rstrip()
    return text
