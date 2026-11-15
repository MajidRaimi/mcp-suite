from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import tzinfo

from office_mcp.features.teams.chats.models import Chat, ChatRow
from office_mcp.features.teams.chats.opener import (
    WRITE_TARGET_HINT,
    ChatOpener,
    refuse_channel,
    resolve_for_write,
)
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.header import (
    HEADER_PREFIX,
    ThreadKind,
    open_thread_id,
    participant_count,
    thread_kind,
)
from office_mcp.features.teams.messages.models import Message
from office_mcp.features.teams.messages.parse import (
    BODY_PREFIX,
    message_pane,
    open_chat_name,
    parse_messages,
)
from office_mcp.shared.errors import (
    ForbiddenActionError,
    TargetError,
    TargetNotFoundError,
    TextEntryError,
)
from office_mcp.shared.node import Node
from office_mcp.shared.waiting import Waiter

MAX_TEXT = 4000
MAX_OLDER = 10
OLDER_WAIT = 4.0
MESSAGE_ID = re.compile(r"\d{1,20}")
CONTENT_PREFIX = "content-"
MEETING_KIND = "meeting"
GROUP_KIND = "group"
MESSAGE_ID_HINT = "use a message id from teams_read"
GROUP_OK_HINT = "pass group_ok=true only when the owner asked to write there"


@dataclass(frozen=True)
class OpenedTarget:
    row: ChatRow
    root: Node
    messages: tuple[Message, ...]
    members: int | None

    @property
    def thread_id(self) -> str:
        return self.row.chat.thread_id

    @property
    def name(self) -> str:
        return self.row.chat.name


def require_text(text: str) -> str:
    if not text.strip():
        raise TextEntryError("the text is empty; nothing was sent")
    if len(text) > MAX_TEXT:
        raise TextEntryError(
            f"the text has {len(text)} characters and the limit is {MAX_TEXT}; nothing was sent"
        )
    return text


def require_message_id(message_id: str) -> str:
    wanted = message_id.strip()
    if not MESSAGE_ID.fullmatch(wanted):
        raise TargetError(MESSAGE_ID_HINT)
    return wanted


def pick_target(query: str, rows: list[ChatRow]) -> ChatRow:
    if not query.strip():
        raise TargetError(WRITE_TARGET_HINT)
    return resolve_for_write(query, rows)


def shown_chat_row(root: Node, thread_id: str) -> ChatRow | None:
    if not thread_id or open_thread_id(root) != thread_id:
        return None
    header = root.find(lambda node: node.dom_id == HEADER_PREFIX + thread_id)
    if header is None:
        return None
    chat = Chat(
        thread_id=thread_id,
        name=open_chat_name(root) or thread_id,
        kind=thread_kind(thread_id).value,
        section="",
        unread=False,
        muted=False,
        presence="",
        preview="",
        last_sender="",
        from_me=False,
        time_label="",
        last_active=None,
    )
    return ChatRow(chat=chat, node=header)


def pick_listed_or_shown(query: str, rows: list[ChatRow], root: Node) -> ChatRow:
    wanted = query.strip()
    refuse_channel(wanted)
    if any(row.chat.thread_id == wanted for row in rows):
        return pick_target(query, rows)
    shown = shown_chat_row(root, wanted)
    return shown if shown is not None else pick_target(query, rows)


def is_meeting(row: ChatRow) -> bool:
    return thread_kind(row.chat.thread_id) is ThreadKind.MEETING or row.chat.kind == MEETING_KIND


def is_group(row: ChatRow) -> bool:
    return thread_kind(row.chat.thread_id) is ThreadKind.GROUP or row.chat.kind == GROUP_KIND


def require_meeting_ok(row: ChatRow, group_ok: bool) -> None:
    if is_meeting(row) and not group_ok:
        raise ForbiddenActionError(
            f"{row.chat.name!r} is a meeting chat that reaches every invitee; {GROUP_OK_HINT}"
        )


def require_size_ok(row: ChatRow, members: int | None, group_ok: bool, limit: int) -> None:
    if group_ok:
        return
    if members is not None and members > limit:
        raise ForbiddenActionError(
            f"{row.chat.name!r} has {members} members, more than the limit of {limit}; "
            + GROUP_OK_HINT
        )
    if members is None and is_group(row):
        raise ForbiddenActionError(
            f"the member count of the group chat {row.chat.name!r} could not be read; "
            + GROUP_OK_HINT
        )


def require_open_thread(root: Node, thread_id: str) -> Node:
    opened = open_thread_id(root)
    if opened != thread_id:
        raise TargetError(
            f"Teams does not show the chat {thread_id} (it shows {opened or 'no chat'}); "
            "nothing was done"
        )
    return root


def message_body(root: Node, message_id: str) -> Node | None:
    wanted = BODY_PREFIX + message_id
    return root.find(lambda node: node.dom_id == wanted)


def require_body(root: Node, message_id: str) -> Node:
    body = message_body(root, message_id)
    if body is None:
        raise TargetNotFoundError(f"the message {message_id} is no longer shown in the chat")
    return body


def find_message(messages: tuple[Message, ...], message_id: str) -> Message | None:
    return next((message for message in messages if message.id == message_id), None)


def first_content_id(root: Node) -> str | None:
    pane = message_pane(root)
    if pane is None:
        return None
    first = pane.find(lambda node: node.dom_id.startswith(CONTENT_PREFIX))
    return first.dom_id if first is not None else None


class ChatTargets:
    def __init__(
        self,
        chats: ChatService,
        opener: ChatOpener,
        gateway: TeamsGateway,
        waiter: Waiter,
        zone: tzinfo,
    ) -> None:
        self._chats = chats
        self._opener = opener
        self._gateway = gateway
        self._waiter = waiter
        self._zone = zone

    def resolve(self, query: str) -> ChatRow:
        return pick_listed_or_shown(query, self._chats.rows(), self._gateway.snapshot())

    def rows(self) -> list[ChatRow]:
        return self._chats.rows()

    def open(self, row: ChatRow) -> OpenedTarget:
        opened = self._opener.open_row(row)
        if opened.thread_id != row.chat.thread_id:
            raise TargetError(f"Teams opened a different chat than {row.chat.name!r}")
        root = self.current(row.chat.thread_id)
        return OpenedTarget(
            row=row,
            root=root,
            messages=self.messages(root),
            members=participant_count(root),
        )

    def current(self, thread_id: str) -> Node:
        root = self._waiter.until(lambda: self._showing(thread_id))
        return require_open_thread(root or self._gateway.snapshot(), thread_id)

    def now(self, thread_id: str) -> Node:
        return require_open_thread(self._gateway.snapshot(), thread_id)

    def messages(self, root: Node) -> tuple[Message, ...]:
        return tuple(parse_messages(root, self._zone))

    def locate(self, opened: OpenedTarget, message_id: str) -> tuple[OpenedTarget, Message]:
        current = opened
        for _ in range(MAX_OLDER + 1):
            found = find_message(current.messages, message_id)
            if found is not None:
                return current, found
            older = self._load_older(current)
            if older is None:
                break
            current = older
        raise TargetNotFoundError(
            f"the message {message_id} is not in the chat {opened.name!r}; {MESSAGE_ID_HINT}"
        )

    def _load_older(self, opened: OpenedTarget) -> OpenedTarget | None:
        first_id = first_content_id(opened.root)
        if first_id is None:
            return None
        first = opened.root.find(lambda node: node.dom_id == first_id)
        if first is None:
            return None
        self._gateway.scroll_into_view(first)
        root = self._waiter.until(lambda: self._moved(opened.thread_id, first_id), OLDER_WAIT)
        if root is None:
            return None
        return OpenedTarget(
            row=opened.row, root=root, messages=self.messages(root), members=opened.members
        )

    def _showing(self, thread_id: str) -> Node | None:
        root = self._gateway.snapshot()
        return root if open_thread_id(root) == thread_id else None

    def _moved(self, thread_id: str, previous_first: str) -> Node | None:
        root = self._gateway.snapshot()
        if open_thread_id(root) != thread_id:
            return None
        first = first_content_id(root)
        return root if first is not None and first != previous_first else None
