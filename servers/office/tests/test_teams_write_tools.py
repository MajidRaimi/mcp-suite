from __future__ import annotations

import anyio
import pytest
from fakes.drivers import fake_runtime, teams_rig
from fakes.lock import MemoryUiLock
from fakes.p5_rig import BOB, CHANNEL, WritingRig, standard_world, writing_rig
from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import teams_window
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.features.teams.writing.factory import build_teams_writing
from office_mcp.features.teams.writing.service import TeamsWritingService
from office_mcp.features.teams.writing.tools import TeamsWritingTools
from office_mcp.shared.drafts import Channel, NoTextGuard
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
    annotations_for,
)
from office_mcp.shared.runner import Lazy, Runner

RULES = "No dashes and no emoji in Teams messages."
DASHES = ("\u2013", "\u2014")


def rules_text(channel: Channel) -> str:
    return RULES if channel is Channel.TEAMS else "other rules"


def tools_for(rig: WritingRig) -> TeamsWritingTools:
    runner = Runner(MemoryUiLock(), rig.clock, lock_timeout=20, stale_after=10)
    return TeamsWritingTools(runner, Lazy(lambda: rig.service), Owner("Alex Example"), rules_text)


def test_specs_cover_tools_31_to_35_with_their_groups_and_effects() -> None:
    specs = {spec.name: spec for spec in tools_for(writing_rig(standard_world())).specs()}

    assert {name: (spec.group, spec.effect, spec.stability) for name, spec in specs.items()} == {
        "teams_send": (ToolGroup.SEND, Effect.SENDS, Stability.STABLE),
        "teams_reply": (ToolGroup.SEND, Effect.SENDS, Stability.STABLE),
        "teams_react": (ToolGroup.SEND, Effect.SENDS, Stability.STABLE),
        "teams_delete_message": (ToolGroup.SEND, Effect.SOFT_DELETES, Stability.STABLE),
        "teams_mark_unread": (ToolGroup.ORGANIZE, Effect.ORGANIZES, Stability.EXPERIMENTAL),
    }
    assert specs["teams_mark_unread"].description.startswith(EXPERIMENTAL_PREFIX)
    for name in ("teams_send", "teams_reply"):
        assert specs[name].description.endswith(
            f"{RULES} Call office_check_text first when unsure."
        )
        assert "Alex Example" in specs[name].description
    for spec in specs.values():
        assert not any(dash in spec.description for dash in DASHES)
        assert spec.function.__name__ == spec.name


def test_tools_register_with_typed_constrained_schemas() -> None:
    server = MCPServer("test")
    for spec in tools_for(writing_rig(standard_world())).specs():
        server.tool(
            name=spec.name, description=spec.description, annotations=annotations_for(spec.effect)
        )(spec.function)

    listed = {tool.name: tool for tool in anyio.run(server.list_tools)}

    send = listed["teams_send"].input_schema
    assert send["required"] == ["chat", "text"]
    assert send["properties"]["text"]["maxLength"] == 4000
    assert send["properties"]["dry_run"]["default"] is False
    react = listed["teams_react"].input_schema
    assert react["properties"]["reaction"]["enum"] == ["like", "heart", "laugh", "surprised"]
    assert listed["teams_send"].annotations is not None
    assert listed["teams_send"].annotations.open_world_hint is True
    assert listed["teams_delete_message"].annotations is not None
    assert listed["teams_delete_message"].annotations.destructive_hint is True


def test_teams_send_returns_the_send_record() -> None:
    rig = writing_rig(standard_world())

    record = tools_for(rig).teams_send("Bob Sample", "On my way.")

    assert record["status"] == "sent"
    assert record["thread_id"] == BOB
    assert record["channel"] == "teams"
    assert rig.world.commits == 1


def test_refusals_become_tool_errors_with_the_reason() -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(ToolError, match="never writes in Teams channels"):
        tools_for(rig).teams_send(CHANNEL, "Hello team", dry_run=True)
    with pytest.raises(ToolError, match="only on messages from others"):
        tools_for(rig).teams_mark_unread("Alex Example (You)")

    assert rig.world.commits == 0


def test_react_delete_and_reply_tools_return_records() -> None:
    rig = writing_rig(standard_world())
    tools = tools_for(rig)

    reacted = tools.teams_react("Bob Sample", "1790000001000", "heart")
    replied = tools.teams_reply("Bob Sample", "1790000001000", "Looks good.")
    deleted = tools.teams_delete_message("Bob Sample", replied["message_ref"])

    assert reacted["state"] == "on"
    assert replied["quoted"]["message_id"] == "1790000001000"
    assert deleted["succeeded"] == 1


def test_build_teams_writing_wires_the_runtime_and_driver() -> None:
    teams = teams_rig(FakeTeamsGateway([teams_window("Chat", [], [])]))

    service = build_teams_writing(fake_runtime(), teams.driver, NoTextGuard())

    assert isinstance(service, TeamsWritingService)
