from __future__ import annotations

import pytest
from fakes.clock import FakeClock
from fakes.drivers import fake_runtime, teams_rig
from fakes.lock import MemoryUiLock
from fakes.p6_chats import OMAR, UnreadDesk
from fakes.teams_gateway import FakeTeamsGateway
from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.features.teams.reading.factory import build_teams_reading
from office_mcp.features.teams.reading.tools import TeamsReadingTools
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import EXPERIMENTAL_PREFIX, Effect, Stability, ToolGroup
from office_mcp.shared.runner import Lazy, Runner

CHAT_KEYS = {
    "thread_id",
    "name",
    "kind",
    "section",
    "unread",
    "muted",
    "presence",
    "preview",
    "last_sender",
    "from_me",
    "time_label",
    "last_active",
}


def tools(gateway: FakeTeamsGateway) -> tuple[TeamsReadingTools, MemoryUiLock]:
    clock = FakeClock()
    lock = MemoryUiLock(clock)
    service = build_teams_reading(fake_runtime(), teams_rig(gateway).driver)
    runner = Runner(lock, clock, lock_timeout=20, stale_after=10)
    return TeamsReadingTools(runner, Lazy(lambda: service), Owner("Jane Doe")), lock


def test_specs_cover_the_seven_read_tools_with_their_effects() -> None:
    specs = tools(FakeTeamsGateway([]))[0].specs()

    assert {spec.name: spec.effect for spec in specs} == {
        "teams_chats": Effect.READS,
        "teams_read": Effect.MARKS_READ,
        "teams_needs_reply": Effect.READS,
        "teams_activity": Effect.MARKS_READ,
        "teams_search": Effect.READS,
        "teams_channels": Effect.MARKS_READ,
        "teams_channel_read": Effect.MARKS_READ,
    }
    assert all(spec.group is ToolGroup.READ for spec in specs)
    assert all(spec.stability is Stability.STABLE for spec in specs)
    assert all(spec.function.__name__ == spec.name for spec in specs)
    assert not any(spec.description.startswith(EXPERIMENTAL_PREFIX) for spec in specs)
    assert "Jane Doe" in next(s for s in specs if s.name == "teams_needs_reply").description


def test_side_effect_disclosures_are_in_the_descriptions() -> None:
    specs = {spec.name: spec.description for spec in tools(FakeTeamsGateway([]))[0].specs()}

    assert "top item" in specs["teams_activity"]
    assert "last channel" in specs["teams_channels"]
    assert "marks it read" in specs["teams_channel_read"]
    assert "marks it read" in specs["teams_read"]
    assert "nothing is marked read" in specs["teams_search"]
    for description in specs.values():
        assert chr(0x2013) not in description and chr(0x2014) not in description


def test_teams_chats_records_keep_their_fields() -> None:
    gateway = FakeTeamsGateway([])
    UnreadDesk(gateway)
    reading, lock = tools(gateway)

    records = reading.teams_chats()

    assert [record["thread_id"] for record in records] == [OMAR, "19:alice@unq.gbl.spaces"]
    assert all(set(record) == CHAT_KEYS for record in records)
    assert lock.acquired == lock.released == 1


def test_teams_read_adds_only_new_fields_and_caps_the_limit() -> None:
    gateway = FakeTeamsGateway([])
    UnreadDesk(gateway)

    record = tools(gateway)[0].teams_read("omar", limit=0, older=99)

    assert set(record) == {"chat", "loaded", "returned", "messages", "restored_unread", "notes"}
    assert record["returned"] == 3
    assert record["restored_unread"] is False
    assert set(record["messages"][0]) == {
        "id",
        "author",
        "mine",
        "sent_at",
        "text",
        "quote",
        "edited",
    }


def test_teams_needs_reply_uses_the_chat_list_only() -> None:
    gateway = FakeTeamsGateway([])
    UnreadDesk(gateway)

    records = tools(gateway)[0].teams_needs_reply(days=7)

    assert [record["name"] for record in records] == ["Omar Haddad", "Alice Example"]
    assert gateway.pressed == []


def test_office_errors_become_tool_errors() -> None:
    gateway = FakeTeamsGateway([])
    UnreadDesk(gateway)

    with pytest.raises(ToolError, match="2 to 200 characters"):
        tools(gateway)[0].teams_search("a")
