from __future__ import annotations

import pytest
from conftest import TODAY
from fakes.clock import fake_waiter
from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import chat_row, message, teams_window

from office_mcp.features.teams.chats.models import ChatRow
from office_mcp.features.teams.chats.opener import ChatOpener, resolve_for_write, shows_thread
from office_mcp.features.teams.chats.parse import parse_chats
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.shared.errors import (
    AmbiguousTargetError,
    ForbiddenActionError,
    TargetError,
    UiTimeoutError,
)
from office_mcp.shared.node import Node

TODAY_ROWS = [
    chat_row("19:bob@unq.gbl.spaces", "Bob Sample", "ok", "10/19"),
    chat_row("19:alice@unq.gbl.spaces", "Alice Example", "hi", "4:09 AM"),
    chat_row("19:sam1@unq.gbl.spaces", "Sam Alpha", "x", "10/18"),
    chat_row("19:sam2@unq.gbl.spaces", "Sam Alpha", "y", "10/17"),
]
BOB = "19:bob@unq.gbl.spaces"


def rows() -> list[ChatRow]:
    return parse_chats(teams_window("Chat", TODAY_ROWS, []), TODAY)


def opener_for(gateway: FakeTeamsGateway) -> ChatOpener:
    waiter = fake_waiter()
    return ChatOpener(gateway, ChatService(gateway, lambda: TODAY, waiter), waiter)


def bob_open() -> Node:
    return teams_window(
        "Chat | Bob Sample",
        TODAY_ROWS,
        message("1790000000000", "Bob Sample", "hello", "October 19, 2026 9:00 AM"),
        thread=BOB,
    )


@pytest.mark.parametrize("query", ["19:general@thread.tacv2", "19:old@thread.skype"])
def test_write_resolution_refuses_channels(query: str) -> None:
    with pytest.raises(ForbiddenActionError):
        resolve_for_write(query, rows())


def test_write_resolution_needs_an_exact_name_or_id() -> None:
    assert resolve_for_write(BOB, rows()).chat.name == "Bob Sample"
    assert resolve_for_write("  bob sample ", rows()).chat.thread_id == BOB
    with pytest.raises(TargetError, match="thread id or the exact chat name"):
        resolve_for_write("Bob", rows())
    with pytest.raises(AmbiguousTargetError):
        resolve_for_write("Sam Alpha", rows())


def test_shows_thread_prefers_the_header_id_over_the_title() -> None:
    bob = resolve_for_write(BOB, rows())
    alice = resolve_for_write("Alice Example", rows())
    wrong_header = teams_window("Chat | Bob Sample", [], [], thread=alice.chat.thread_id)

    assert shows_thread(bob_open(), bob)
    assert not shows_thread(wrong_header, bob)


def test_opener_presses_the_row_and_waits_for_the_thread() -> None:
    gateway = FakeTeamsGateway([teams_window("Chat", TODAY_ROWS, [])])
    gateway.on_press = lambda node: gateway.show(bob_open())

    opened = opener_for(gateway).open("Bob Sample", for_write=True)

    assert opened.thread_id == BOB
    assert gateway.pressed == [BOB]
    assert open_title(opened.root) == "Chat | Bob Sample"


def test_opener_skips_the_press_when_the_chat_is_already_open() -> None:
    gateway = FakeTeamsGateway([bob_open()])

    opened = opener_for(gateway).open(BOB, for_write=False)

    assert opened.row.chat.name == "Bob Sample"
    assert gateway.pressed == []


def test_opener_for_write_refuses_a_partial_name_before_any_press() -> None:
    gateway = FakeTeamsGateway([teams_window("Chat", TODAY_ROWS, [])])

    with pytest.raises(TargetError):
        opener_for(gateway).open("Bob", for_write=True)
    assert gateway.pressed == []


def test_opener_times_out_when_another_chat_opens() -> None:
    alice = teams_window("Chat | Alice Example", TODAY_ROWS, [], thread="19:alice@unq.gbl.spaces")
    gateway = FakeTeamsGateway([teams_window("Chat", TODAY_ROWS, [])])
    gateway.on_press = lambda node: gateway.show(alice)

    with pytest.raises(UiTimeoutError):
        opener_for(gateway).open(BOB, for_write=True)


def test_find_self_looks_up_a_row_by_thread_id() -> None:
    opener = opener_for(FakeTeamsGateway([teams_window("Chat", TODAY_ROWS, [])]))

    found = opener.find_self(BOB)

    assert found is not None
    assert found.chat.name == "Bob Sample"
    assert opener.find_self("19:missing@unq.gbl.spaces") is None


def open_title(root: Node) -> str:
    return root.title
