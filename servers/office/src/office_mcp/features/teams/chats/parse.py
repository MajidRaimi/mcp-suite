from __future__ import annotations

import re
from collections.abc import Callable, Iterator
from datetime import date

from office_mcp.features.teams.chats.models import Chat, ChatRow
from office_mcp.shared.clock import list_label_date
from office_mcp.shared.node import Node

ROW_ROLE = "AXRow"
TITLE_PREFIX = "title-chat-list-item_"
TIME_PREFIX = "time-chat-list-item_"
PREVIEW_PREFIX = "message-preview-chat-list-item_"
UNREAD_LABEL = "unread-label"
UNREAD_TITLE = "Unread message "
KINDS = (("Group chat ", "group"), ("Meeting chat ", "meeting"), ("Chat ", "chat"))
LAST_MESSAGE = " Last message"
MUTED = re.compile(r"\bMuted\b")
FROM_ME = "You:"


def parse_chats(root: Node, today: date) -> list[ChatRow]:
    rows: list[ChatRow] = []
    seen: set[str] = set()
    section = ""
    for node in root.walk():
        if node.role != ROW_ROLE:
            continue
        title_node = own_find(node, lambda child: child.dom_id.startswith(TITLE_PREFIX))
        if title_node is None:
            section = node.label()
            continue
        thread_id = title_node.dom_id.removeprefix(TITLE_PREFIX)
        if thread_id in seen:
            continue
        seen.add(thread_id)
        rows.append(
            ChatRow(chat=build_chat(node, title_node, thread_id, section, today), node=node)
        )
    return rows


def build_chat(row: Node, title_node: Node, thread_id: str, section: str, today: date) -> Chat:
    name = title_node.text()
    title = row.title or row.description
    kind = kind_of(title)
    preview = prefixed_text(row, PREVIEW_PREFIX)
    time_label = prefixed_text(row, TIME_PREFIX)
    from_me = preview.startswith(FROM_ME)
    active = list_label_date(time_label, today)
    return Chat(
        thread_id=thread_id,
        name=name,
        kind=kind,
        section=section,
        unread=title.startswith(UNREAD_TITLE) or own_find(row, is_unread_label) is not None,
        muted=MUTED.search(title) is not None,
        presence=presence_of(title, name) if kind == "chat" else "",
        preview=preview,
        last_sender=sender_of(preview, name, kind, from_me),
        from_me=from_me,
        time_label=time_label,
        last_active=active.isoformat() if active else None,
    )


def own_nodes(row: Node) -> Iterator[Node]:
    for child in row.children:
        if child.role == ROW_ROLE:
            continue
        yield child
        yield from own_nodes(child)


def own_find(row: Node, predicate: Callable[[Node], bool]) -> Node | None:
    return next((node for node in own_nodes(row) if predicate(node)), None)


def is_unread_label(node: Node) -> bool:
    return node.dom_id == UNREAD_LABEL


def prefixed_text(row: Node, prefix: str) -> str:
    node = own_find(row, lambda child: child.dom_id.startswith(prefix))
    return node.text() if node else ""


def kind_of(title: str) -> str:
    stripped = title.removeprefix(UNREAD_TITLE)
    return next((kind for prefix, kind in KINDS if stripped.startswith(prefix)), "chat")


def presence_of(title: str, name: str) -> str:
    _, found, after = title.partition(name)
    if not found or LAST_MESSAGE not in after:
        return ""
    return after.split(LAST_MESSAGE, 1)[0].strip()


def sender_of(preview: str, name: str, kind: str, from_me: bool) -> str:
    if from_me:
        return "you"
    if kind == "chat":
        return name
    sender, separator, _ = preview.partition(":")
    return sender.strip() if separator and len(sender) <= 60 else ""
