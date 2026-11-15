from __future__ import annotations

from datetime import UTC, datetime, tzinfo
from zoneinfo import ZoneInfo

import pytest
from conftest import TODAY
from fakes.clock import FakeClock, fake_waiter
from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import chat_row, message, teams_window

from office_mcp.features.teams.chats.opener import ChatOpener
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.messages.parse import parse_messages
from office_mcp.features.teams.messages.service import MessageService
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.node import Node

ROWS = [
    chat_row("19:b", "Bob Sample", "You: yes", "10/19"),
    chat_row("19:a", "Alice Example", "ok", "4:09 AM", unread=True),
]
FIRST = "1790000000000"
SECOND = "1790000001000"
THIRD = "1790000011000"


def bob_messages() -> list[Node]:
    return [
        *message(
            FIRST, "Jane Doe", "in a meeting, back soon", "September 21, 2026 2:13 PM", mine=True
        ),
        *message(SECOND, "Bob Sample", "مرحبا", "September 21, 2026 2:13 PM"),
        *message(THIRD, "Bob Sample", "شكرا", "September 21, 2026 2:13 PM", show_author=False),
    ]


def message_service(
    gateway: FakeTeamsGateway, timeout: float = 12.0, zone: tzinfo = UTC
) -> MessageService:
    waiter = fake_waiter(FakeClock(), timeout=timeout)
    opener = ChatOpener(gateway, ChatService(gateway, lambda: TODAY, waiter), waiter)
    return MessageService(gateway, opener, zone, waiter)


def test_parses_authors_text_ownership_and_time() -> None:
    messages = parse_messages(teams_window("Chat | Bob Sample", ROWS, bob_messages()), UTC)

    assert [(m.author, m.mine, m.text) for m in messages] == [
        ("Jane Doe", True, "in a meeting, back soon"),
        ("Bob Sample", False, "مرحبا"),
        ("Bob Sample", False, "شكرا"),
    ]
    assert messages[0].sent_at == "2026-09-21T14:13:20+00:00"
    sent = datetime.fromisoformat(messages[0].sent_at)
    assert sent == datetime(2026, 9, 21, 14, 13, 20, tzinfo=UTC)


def test_message_times_follow_the_injected_zone() -> None:
    zone = ZoneInfo("Europe/Istanbul")
    messages = parse_messages(teams_window("Chat | Bob Sample", ROWS, bob_messages()), zone)

    assert messages[0].sent_at == "2026-09-21T17:13:20+03:00"


def test_message_without_content_node_strips_author_and_stamp() -> None:
    card = Node(
        role="AXGroup",
        title="Alert fired Workflows September 21, 2026 2:13 PM.",
        dom_id="message-body-card1",
        classes=("fui-ChatMessage__body",),
    )
    messages = parse_messages(
        teams_window(
            "Chat | Ops Alerts",
            [],
            [Node(role="AXHeading", title="Alert fired by Workflows", value="4"), card],
        ),
        UTC,
    )

    assert messages[0].author == "Workflows"
    assert messages[0].text == "Alert fired Workflows September 21, 2026 2:13 PM."
    assert messages[0].sent_at == ""


def test_read_opens_the_chat_when_another_is_showing() -> None:
    before = teams_window("Chat | Alice Example", ROWS, [])
    after = teams_window("Chat | Bob Sample", ROWS, bob_messages())
    gateway = FakeTeamsGateway([before, before])
    gateway.on_press = lambda _node: gateway.screens.extend([after])

    conversation = message_service(gateway).read("bob", limit=2, older=0)

    assert gateway.pressed == ["19:b"]
    assert conversation.chat.name == "Bob Sample"
    assert conversation.loaded == 3
    assert [m.text for m in conversation.messages] == ["مرحبا", "شكرا"]


def test_read_does_not_click_when_chat_is_already_open() -> None:
    gateway = FakeTeamsGateway([teams_window("Chat | Bob Sample", ROWS, bob_messages())])

    conversation = message_service(gateway).read("bob", limit=0, older=0)

    assert gateway.pressed == []
    assert conversation.loaded == 3


def test_read_trusts_the_header_thread_id_over_the_title() -> None:
    stale_title = teams_window("Chat | Bob Sample", ROWS, bob_messages(), thread="19:a")
    opened = teams_window("Chat", ROWS, bob_messages(), thread="19:b")
    gateway = FakeTeamsGateway([stale_title, stale_title])
    gateway.on_press = lambda _node: gateway.screens.extend([opened])

    conversation = message_service(gateway).read("bob", limit=0, older=0)

    assert gateway.pressed == ["19:b"]
    assert conversation.chat.thread_id == "19:b"


def test_read_loads_older_messages_by_scrolling() -> None:
    newer = teams_window("Chat | Bob Sample", ROWS, bob_messages()[3:])
    older = teams_window("Chat | Bob Sample", ROWS, bob_messages())
    gateway = FakeTeamsGateway([newer, newer, newer, older])

    conversation = message_service(gateway).read("bob", limit=0, older=3)

    assert gateway.scrolled[0] == f"content-{SECOND}"
    assert conversation.loaded == 3


def test_read_times_out_when_chat_never_opens() -> None:
    stuck = teams_window("Chat | Alice Example", ROWS, [])
    gateway = FakeTeamsGateway([stuck])

    with pytest.raises(OfficeError, match="did not open"):
        message_service(gateway, timeout=2).read("bob", limit=5, older=0)
