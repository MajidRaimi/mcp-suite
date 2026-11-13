from __future__ import annotations

import pytest
from conftest import TODAY, FakeGateway, chat_row, section_row, window

from office_mcp.features.teams.chats.parse import parse_chats
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.replies.service import needs_reply
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.node import Node


def chat_list() -> Node:
    rows = [
        section_row("Quick views"),
        section_row("Drafts"),
        section_row(
            "Pinned ⭐",
            [chat_row("19:self", "Jane Doe (You)", "You: note", "6/17", presence="")],
        ),
        section_row(
            "Project Group",
            [
                chat_row("19:a", "Alice Example", "تم", "4:09 AM", presence="Away", unread=True),
                chat_row("19:b", "Bob Sample", "You: sounds good, go ahead", "10/19"),
                chat_row(
                    "19:c", "Carol Test", "Can you review this?", "10/11", presence="In a call"
                ),
            ],
        ),
        section_row(
            "Chats",
            [
                chat_row(
                    "19:g@thread.v2",
                    "Ops Alerts",
                    "Workflows: Sent a card",
                    "6:32 AM",
                    kind="Group chat",
                    muted=True,
                ),
                chat_row(
                    "19:h@thread.v2",
                    "Team Chat",
                    "Erin: Useful reference",
                    "Yesterday",
                    kind="Group chat",
                    unread=True,
                ),
                chat_row(
                    "19:m@thread.v2",
                    "Project Sync Team",
                    "Dan: latest updates",
                    "10/18",
                    kind="Meeting chat",
                ),
                chat_row("19:b", "Bob Sample", "You: sounds good, go ahead", "10/19"),
            ],
        ),
    ]
    return window("Chat | Bob Sample", rows, [])


def test_parses_rows_with_sections_kinds_and_flags() -> None:
    chats = [row.chat for row in parse_chats(chat_list(), TODAY)]

    assert [chat.name for chat in chats] == [
        "Jane Doe (You)",
        "Alice Example",
        "Bob Sample",
        "Carol Test",
        "Ops Alerts",
        "Team Chat",
        "Project Sync Team",
    ]
    assert chats[0].section == "Pinned ⭐"
    alice, bob, carol = chats[1], chats[2], chats[3]
    assert alice.unread and alice.presence == "Away" and alice.section == "Project Group"
    assert alice.last_sender == "Alice Example" and not alice.from_me
    assert alice.last_active == "2026-10-20"
    assert bob.from_me and bob.last_sender == "you" and not bob.unread
    assert carol.presence == "In a call" and carol.last_active == "2026-10-11"
    alerts, team, meeting = chats[4], chats[5], chats[6]
    assert alerts.kind == "group" and alerts.muted and alerts.presence == ""
    assert team.unread and team.last_sender == "Erin" and team.last_active == "2026-10-19"
    assert meeting.kind == "meeting" and meeting.section == "Chats"


def test_list_filters_unread_and_section() -> None:
    service = ChatService(FakeGateway([chat_list()]), lambda: TODAY)

    assert [c.name for c in service.filtered(True, None, 0)] == ["Alice Example", "Team Chat"]
    assert [c.name for c in service.filtered(False, "pinned", 0)] == ["Jane Doe (You)"]
    assert len(service.filtered(False, None, 2)) == 2


def test_resolve_prefers_exact_then_prefix_then_words() -> None:
    service = ChatService(FakeGateway([chat_list()]), lambda: TODAY)
    rows = service.rows()

    assert service.resolve("bob", rows).chat.thread_id == "19:b"
    assert service.resolve("alice example", rows).chat.thread_id == "19:a"
    assert service.resolve("sync project", rows).chat.thread_id == "19:m@thread.v2"
    assert service.resolve("19:h@thread.v2", rows).chat.name == "Team Chat"
    with pytest.raises(OfficeError, match="no chat named"):
        service.resolve("nobody here", rows)


def test_resolve_reports_ambiguous_matches() -> None:
    rows = parse_chats(
        window(
            "Chat | x",
            [
                chat_row("19:1", "Sam Alpha", "hi", "10/19"),
                chat_row("19:2", "Sam Beta", "yo", "10/19"),
            ],
            [],
        ),
        TODAY,
    )
    service = ChatService(FakeGateway([window("Chat | x", [], [])]), lambda: TODAY)

    with pytest.raises(OfficeError, match="Sam Alpha, Sam Beta"):
        service.resolve("sam", rows)
    assert service.resolve("alpha", rows).chat.thread_id == "19:1"


def test_switches_to_chat_app_when_list_is_hidden() -> None:
    activity = Node(
        role="AXWindow",
        title="Activity | Feed",
        children=(Node(role="AXCheckBox", title="Chat (⌘ 2)", handle="chat-app"),),
    )
    gateway = FakeGateway([activity, chat_list()])

    rows = ChatService(gateway, lambda: TODAY).rows()

    assert gateway.pressed == ["chat-app"]
    assert len(rows) == 7


def test_needs_reply_skips_mine_muted_self_and_old() -> None:
    chats = [row.chat for row in parse_chats(chat_list(), TODAY)]

    waiting = needs_reply(chats, TODAY, days=7, include_meetings=False)

    assert [chat.name for chat in waiting] == ["Alice Example", "Team Chat"]
    with_meetings = needs_reply(chats, TODAY, days=7, include_meetings=True)
    assert [chat.name for chat in with_meetings][-1] == "Project Sync Team"
    assert [c.name for c in needs_reply(chats, TODAY, 30, False)][-1] == "Carol Test"
