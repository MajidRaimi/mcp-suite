from __future__ import annotations

import pytest
from conftest import TODAY
from fakes.clock import FakeClock
from fakes.trees import chat_row, teams_window

from office_mcp.features.teams.chats.models import ChatRow
from office_mcp.features.teams.chats.parse import parse_chats
from office_mcp.features.teams.messages.models import Message
from office_mcp.features.teams.writing.compose import strip_quote
from office_mcp.features.teams.writing.guard import TeamsSendGuard
from office_mcp.features.teams.writing.targets import (
    pick_target,
    require_message_id,
    require_size_ok,
)
from office_mcp.features.teams.writing.verify import (
    find_sent,
    newest_from_others,
    newest_id,
    recent_duplicate,
    sent_early,
)
from office_mcp.shared.errors import (
    AmbiguousTargetError,
    DuplicateSendError,
    ForbiddenActionError,
    TargetError,
    TextEntryError,
)
from office_mcp.shared.node import Node

GROUP = "19:atlasteam@thread.v2"


def rows(*nodes: Node) -> list[ChatRow]:
    return parse_chats(teams_window("Chat", list(nodes), []), TODAY)


def mine(message_id: str, text: str) -> Message:
    return Message(id=message_id, author="Alex Example", mine=True, sent_at="", text=text)


def theirs(message_id: str, text: str, author: str = "Bob Sample") -> Message:
    return Message(id=message_id, author=author, mine=False, sent_at="", text=text)


def test_pick_target_needs_the_full_name_and_a_unique_match() -> None:
    listed = rows(
        chat_row("19:a@unq.gbl.spaces", "Jane Doe", "hi", "9:00 AM"),
        chat_row("19:b@unq.gbl.spaces", "Jane Doe", "yo", "8:00 AM"),
        chat_row("19:c@unq.gbl.spaces", "Omar Haddad", "ok", "7:00 AM"),
    )

    assert pick_target(" omar haddad ", listed).chat.thread_id == "19:c@unq.gbl.spaces"
    assert pick_target("19:a@unq.gbl.spaces", listed).chat.name == "Jane Doe"
    with pytest.raises(AmbiguousTargetError):
        pick_target("Jane Doe", listed)
    for partial in ("Omar", "Haddad", "omar h", ""):
        with pytest.raises(TargetError):
            pick_target(partial, listed)


def test_message_ids_are_digits_only() -> None:
    assert require_message_id(" 1790000001000 ") == "1790000001000"
    for bad in ("", "abc", "17900-1", "1" * 21):
        with pytest.raises(TargetError):
            require_message_id(bad)


def test_group_size_rule_fails_closed_for_an_unknown_group_count() -> None:
    group = rows(chat_row(GROUP, "Project Atlas team", "x", "9:00 AM", kind="Group chat"))[0]
    single = rows(chat_row("19:a@unq.gbl.spaces", "Jane Doe", "x", "9:00 AM"))[0]

    require_size_ok(single, None, group_ok=False, limit=25)
    require_size_ok(group, 25, group_ok=False, limit=25)
    require_size_ok(group, None, group_ok=True, limit=25)
    with pytest.raises(ForbiddenActionError, match="could not be read"):
        require_size_ok(group, None, group_ok=False, limit=25)
    with pytest.raises(ForbiddenActionError, match="26 members"):
        require_size_ok(group, 26, group_ok=False, limit=25)


def test_verification_matches_a_newer_own_message_by_id_and_text() -> None:
    before = [theirs("100", "Question?"), mine("200", "Earlier")]
    after = [*before, mine("150", "Hello  there"), mine("300", "Hello there")]

    assert newest_id(before) == 200
    assert newest_id([theirs("500", "Later"), mine("200", "Mine")]) == 500
    assert find_sent(after, 200, "Hello\nthere") is after[-1]
    assert find_sent(after, 300, "Hello there") is None
    assert find_sent([*before, mine("300", "Other")], 200, "Hello there") is None
    quoted = mine("400", "Begin quote, Bob Sample, 10/20/2026, 9:00 AM, Question? Yes")
    assert find_sent([quoted], 200, "Yes", quoted=True) is quoted
    assert find_sent([quoted], 200, "Yes") is None


def test_early_send_is_any_newer_own_message() -> None:
    messages = [mine("200", "Earlier"), theirs("250", "Reply")]

    assert not sent_early(messages, 200)
    assert sent_early([*messages, mine("260", "First line")], 200)


def test_context_and_duplicates_use_the_right_messages() -> None:
    messages = [theirs("1", "Old", "Jane Doe"), theirs("2", "Newest"), mine("3", "Sure, Thing")]

    assert newest_from_others(messages) is messages[1]
    assert newest_from_others([mine("3", "x")]) is None
    assert recent_duplicate(messages, "sure,   thing") is messages[2]
    many = [mine(str(number), f"Note {number}") for number in range(10)]
    assert recent_duplicate(many, "Note 9") is not None
    assert recent_duplicate(many, "Note 2") is None


def test_send_guard_blocks_within_the_window_only() -> None:
    clock = FakeClock()
    guard = TeamsSendGuard(clock)

    guard.record("48:notes", "Hello  there")
    with pytest.raises(DuplicateSendError):
        guard.check("48:notes", "hello there")
    guard.check("19:other@unq.gbl.spaces", "Hello there")
    clock.advance(121)
    guard.check("48:notes", "Hello there")


def test_strip_quote_requires_the_quote_to_stay_in_front() -> None:
    assert strip_quote("Bob: Question?\nYes", "Bob: Question?\n") == "Yes"
    assert strip_quote("Yes", "") == "Yes"
    with pytest.raises(TextEntryError):
        strip_quote("Yes", "Bob: Question?\n")
