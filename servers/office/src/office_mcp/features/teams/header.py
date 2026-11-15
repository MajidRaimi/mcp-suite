from __future__ import annotations

import re
from enum import StrEnum

from office_mcp.shared.node import Node

HEADER_PREFIX = "chat-header-"
PERSON_PREFIX = "chat-topic-person-"
PARTICIPANTS_LABEL = "Chat participants"
LIST_ROLE = "AXList"
OVERFLOW = re.compile(r"^\+\s*(?P<count>\d+)")
CHANNEL_MARKS = ("@thread.tacv2", "@thread.skype")
MEETING_PREFIX = "19:meeting_"
GROUP_MARK = "@thread.v2"
NOTES_PREFIX = "48:"


class ThreadKind(StrEnum):
    CHAT = "chat"
    GROUP = "group"
    MEETING = "meeting"
    CHANNEL = "channel"
    NOTES = "notes"


def open_thread_id(root: Node) -> str | None:
    header = root.find(lambda node: node.dom_id.startswith(HEADER_PREFIX))
    if header is None:
        return None
    thread_id = header.dom_id.removeprefix(HEADER_PREFIX).strip()
    return thread_id or None


def participant_count(root: Node) -> int | None:
    roster = root.find(
        lambda node: node.role == LIST_ROLE and PARTICIPANTS_LABEL in (node.title, node.description)
    )
    if roster is not None:
        return roster_count(roster)
    if thread_kind(open_thread_id(root) or "") is not ThreadKind.CHAT:
        return None
    people = {node.dom_id for node in root.walk() if node.dom_id.startswith(PERSON_PREFIX)}
    return len(people) or None


def roster_count(roster: Node) -> int:
    count = 0
    for child in roster.children:
        overflow = OVERFLOW.match(child.label())
        count += int(overflow["count"]) if overflow else 1
    return count


def thread_kind(thread_id: str) -> ThreadKind:
    lowered = thread_id.strip().casefold()
    if lowered.startswith(NOTES_PREFIX):
        return ThreadKind.NOTES
    if any(mark in lowered for mark in CHANNEL_MARKS):
        return ThreadKind.CHANNEL
    if lowered.startswith(MEETING_PREFIX):
        return ThreadKind.MEETING
    if GROUP_MARK in lowered:
        return ThreadKind.GROUP
    return ThreadKind.CHAT
