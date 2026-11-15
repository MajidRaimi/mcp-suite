from __future__ import annotations

import re
from dataclasses import replace
from datetime import date

from office_mcp.features.teams.search.models import ResultType, SearchResult
from office_mcp.shared.clock import list_label_date
from office_mcp.shared.node import STATIC_TEXT, Node

INPUT_ID = "ms-searchux-input"
POPUP_PREFIX = "ms-searchux-popup-"
MESSAGE = re.compile(
    r"^Message from (?P<sender>.+?) at (?P<where>.+?): (?P<snippet>.*)$", re.DOTALL
)
MESSAGE_PREFIX = "Message from "
GROUP_CHAT_PREFIX = "Group chat "
PERSON_PREFIX = "Person "
SUGGESTION_PREFIX = "Suggestion "
FILE_MARK = " file result "
FIELD_SEPARATOR = ", "
IGNORED = frozenset(
    {
        "messages",
        "files",
        "channels",
        "group chats",
        "meetings",
        "images",
        "people",
        "recent messages",
        "clear query",
    }
)
IGNORED_PREFIX = "find in this chat"


def search_input(root: Node) -> Node | None:
    return root.find(lambda node: node.dom_id == INPUT_ID)


def search_popup(root: Node) -> Node | None:
    return root.find(lambda node: node.dom_id.startswith(POPUP_PREFIX))


def popup_texts(popup: Node) -> list[str]:
    return [
        " ".join(node.value.split())
        for node in popup.walk()
        if node.role == STATIC_TEXT and node.value.strip()
    ]


def parse_results(root: Node, today: date) -> list[SearchResult]:
    popup = search_popup(root)
    if popup is None:
        return []
    results: list[SearchResult] = []
    extras_open = False
    for text in popup_texts(popup):
        if is_ignored(text):
            extras_open = False
            continue
        result = typed_result(text)
        if result is not None:
            results.append(result)
            extras_open = result.type is ResultType.MESSAGE
        elif extras_open:
            results[-1] = with_extra(results[-1], text, today)
    return results


def is_ignored(text: str) -> bool:
    folded = text.casefold()
    return folded in IGNORED or folded.startswith(IGNORED_PREFIX)


def typed_result(text: str) -> SearchResult | None:
    message = MESSAGE.match(text)
    if message:
        snippet = message["snippet"].strip()
        return SearchResult(type=ResultType.MESSAGE, text=f"{message['sender']}: {snippet}")
    if text.startswith(MESSAGE_PREFIX):
        return SearchResult(type=ResultType.MESSAGE, text=text.removeprefix(MESSAGE_PREFIX))
    if text.startswith(GROUP_CHAT_PREFIX):
        name, _, members = text.removeprefix(GROUP_CHAT_PREFIX).partition(FIELD_SEPARATOR)
        return SearchResult(type=ResultType.GROUP_CHAT, text=name, chat=name, members=members)
    if text.startswith(PERSON_PREFIX):
        return SearchResult(type=ResultType.PERSON, text=text.removeprefix(PERSON_PREFIX))
    if text.startswith(SUGGESTION_PREFIX):
        return SearchResult(type=ResultType.SUGGESTION, text=text.removeprefix(SUGGESTION_PREFIX))
    if FILE_MARK in text.casefold():
        start = text.casefold().index(FILE_MARK) + len(FILE_MARK)
        return SearchResult(type=ResultType.FILE, text=text[start:].strip())
    return None


def with_extra(result: SearchResult, text: str, today: date) -> SearchResult:
    if not result.time and list_label_date(text, today) is not None:
        return replace(result, time=text)
    if not result.chat:
        return replace(result, chat=text)
    if not result.members:
        return replace(result, members=text)
    if not result.time:
        return replace(result, time=text)
    return result
