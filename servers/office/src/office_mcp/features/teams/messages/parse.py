from __future__ import annotations

import re
from collections.abc import Callable, Iterator
from dataclasses import replace
from datetime import tzinfo

from office_mcp.features.teams.messages.models import Message, Quote
from office_mcp.shared.clock import message_time
from office_mcp.shared.node import STATIC_TEXT, Node

PANE_ID = "chat-pane-list"
BODY_PREFIX = "message-body-"
AUTHOR_PREFIX = "author-"
TIME_PREFIX = "timestamp-"
CONTENT_PREFIX = "content-"
EDITED_PREFIX = "edited-"
EDITED_LABEL = "edited"
MY_BODY_CLASS = "fui-ChatMyMessage__body"
HEADING_ROLE = "AXHeading"
MESSAGE_HEADING_LEVEL = "4"
BY = " by "
WINDOW_SEPARATOR = " | "
QUOTE_PREFIX = "Begin quote, "
QUOTE = re.compile(
    r"^Begin quote, (?P<author>.+?), (?P<date>\d{1,2}/\d{1,2}/\d{2,4}),? "
    r"(?P<time>\d{1,2}:\d{2}(?:\s?[AaPp]\.?[Mm]\.?)?),?\s*(?P<text>.*)$",
    re.DOTALL,
)
QUOTE_FALLBACK_FIELDS = 3


def message_pane(root: Node) -> Node | None:
    return root.find(lambda node: node.dom_id == PANE_ID)


def open_chat_name(root: Node) -> str:
    _, separator, name = root.title.partition(WINDOW_SEPARATOR)
    return name.strip() if separator else ""


def first_content(root: Node) -> Node | None:
    pane = message_pane(root)
    return pane.find(lambda node: node.dom_id.startswith(CONTENT_PREFIX)) if pane else None


def message_body(root: Node, message_id: str) -> Node | None:
    return root.find(lambda node: node.dom_id == BODY_PREFIX + message_id)


def is_body(node: Node) -> bool:
    return node.dom_id.startswith(BODY_PREFIX)


def is_message_heading(node: Node) -> bool:
    return node.role == HEADING_ROLE and node.value == MESSAGE_HEADING_LEVEL and BY in node.title


def is_edited_mark(node: Node) -> bool:
    if node.dom_id.startswith(EDITED_PREFIX):
        return True
    shown = node.value if node.role == STATIC_TEXT else own_title(node)
    return shown.strip().casefold() == EDITED_LABEL


def walk_until(node: Node, stop: Callable[[Node], bool]) -> Iterator[Node]:
    yield node
    if stop(node):
        return
    for child in node.children:
        yield from walk_until(child, stop)


def walk_outside(node: Node, excluded: Node) -> Iterator[Node]:
    if node is excluded:
        return
    yield node
    for child in node.children:
        yield from walk_outside(child, excluded)


def parse_messages(root: Node, zone: tzinfo) -> list[Message]:
    pane = message_pane(root)
    if pane is None:
        return []
    by_id = {node.dom_id: node for node in pane.walk() if node.dom_id}
    marks = EditedMarks.of(pane)
    messages: list[Message] = []
    heading_author = ""
    pending = False
    after_body = False
    for node in outside_bodies(pane):
        if is_message_heading(node):
            heading_author = node.title.rsplit(BY, 1)[1].strip()
            pending = after_body = False
        elif is_body(node):
            message_id = node.dom_id.removeprefix(BODY_PREFIX)
            previous = messages[-1].author if messages else ""
            message = build_message(message_id, node, by_id, heading_author or previous, zone)
            if pending or marks.covers(message_id) or node.find(is_edited_mark) is not None:
                message = replace(message, edited=True)
            messages.append(message)
            heading_author = ""
            pending = False
            after_body = True
        elif is_edited_mark(node) and not marks.placed(node):
            if after_body and messages:
                messages[-1] = replace(messages[-1], edited=True)
            else:
                pending = True
    return messages


def outside_bodies(pane: Node) -> Iterator[Node]:
    return walk_until(pane, lambda candidate: candidate is not pane and is_body(candidate))


class EditedMarks:
    def __init__(self, ids: set[str], nodes: set[int]) -> None:
        self._ids = ids
        self._nodes = nodes

    @classmethod
    def of(cls, pane: Node) -> EditedMarks:
        ids: set[str] = set()
        nodes: set[int] = set()
        for node in outside_bodies(pane):
            if node.dom_id.startswith(EDITED_PREFIX):
                ids.add(node.dom_id.removeprefix(EDITED_PREFIX))
                nodes.add(id(node))
            group_marks = [
                child for child in node.children if is_edited_mark(child) and not is_body(child)
            ]
            owners = header_ids(node)
            if group_marks and len(owners) == 1:
                ids.update(owners)
                nodes.update(id(mark) for mark in group_marks)
        return cls(ids, nodes)

    def covers(self, message_id: str) -> bool:
        return message_id in self._ids

    def placed(self, node: Node) -> bool:
        return id(node) in self._nodes


def header_ids(group: Node) -> set[str]:
    return {
        node.dom_id.removeprefix(prefix)
        for node in outside_bodies(group)
        for prefix in (AUTHOR_PREFIX, TIME_PREFIX)
        if node.dom_id.startswith(prefix)
    }


def build_message(
    message_id: str, body: Node, by_id: dict[str, Node], fallback: str, zone: tzinfo
) -> Message:
    author_node = by_id.get(AUTHOR_PREFIX + message_id)
    stamp_node = by_id.get(TIME_PREFIX + message_id)
    content_node = by_id.get(CONTENT_PREFIX + message_id)
    author = author_node.text() if author_node else fallback
    stamp = stamp_node.label().rstrip(".") if stamp_node else ""
    sent = message_time(message_id, zone)
    quote_node = find_quote(content_node or body)
    return Message(
        id=message_id,
        author=author,
        mine=body.has_class(MY_BODY_CLASS),
        sent_at=sent.isoformat(timespec="seconds") if sent else stamp,
        text=message_text(body, content_node, author, stamp, quote_node),
        quote=parse_quote(quote_node) if quote_node is not None else None,
    )


def message_text(
    body: Node, content: Node | None, author: str, stamp: str, quote: Node | None = None
) -> str:
    if content is not None:
        return reply_text(content, quote) if quote is not None else content.label()
    text = body.label()
    suffixes = (f"{author} {stamp}.", f"{author} {stamp}") if stamp else (author,)
    for suffix in suffixes:
        text = text.removesuffix(suffix).rstrip()
    if quote is None:
        return text
    return without_quote(text, body, quote)


def own_title(node: Node) -> str:
    return " ".join((node.title or node.description).split())


def find_quote(scope: Node) -> Node | None:
    candidates = [
        node
        for node in scope.walk()
        if node is not scope and own_title(node).startswith(QUOTE_PREFIX)
    ]
    return min(candidates, key=lambda node: len(own_title(node))) if candidates else None


def parse_quote(node: Node) -> Quote:
    match = QUOTE.match(own_title(node))
    if match:
        return Quote(
            author=match["author"].strip(),
            sent=f"{match['date']} {match['time']}",
            text=match["text"].strip(),
        )
    pieces = static_pieces(node.walk())
    if len(pieces) >= QUOTE_FALLBACK_FIELDS:
        return Quote(author=pieces[0], sent=pieces[1], text=" ".join(pieces[2:]))
    return Quote(author="", sent="", text=" ".join(pieces))


def reply_text(container: Node, quote: Node) -> str:
    return without_quote(container.label(), container, quote)


def without_quote(text: str, container: Node, quote: Node) -> str:
    quoted = own_title(quote)
    if quoted and quoted in text:
        return " ".join(text.replace(quoted, " ", 1).split())
    return " ".join(static_pieces(walk_outside(container, quote)))


def static_pieces(nodes: Iterator[Node]) -> list[str]:
    return [
        " ".join(node.value.split())
        for node in nodes
        if node.role == STATIC_TEXT and node.value.strip()
    ]
