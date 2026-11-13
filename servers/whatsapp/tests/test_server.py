from __future__ import annotations

from pathlib import Path

import anyio
from mcp.types import Tool
from whatsapp_fakes import FakeWaha, make_http, make_settings

from whatsapp_mcp.server import build_server
from whatsapp_mcp.shared.settings import Settings

READ_ONLY_TOOLS = {
    "whatsapp_needs_reply",
    "whatsapp_recent_chats",
    "whatsapp_read_chat",
    "whatsapp_search",
    "whatsapp_status",
}


def _tools(settings: Settings) -> dict[str, Tool]:
    server = build_server(settings, make_http(FakeWaha()))
    tools = anyio.run(server.list_tools)
    return {tool.name: tool for tool in tools}


def test_every_tool_has_complete_annotations(tmp_path: Path) -> None:
    tools = _tools(make_settings(tmp_path))
    assert set(tools) == {*READ_ONLY_TOOLS, "whatsapp_draft_reply", "whatsapp_send_draft"}
    for name, tool in tools.items():
        annotations = tool.annotations
        assert annotations is not None
        assert annotations.read_only_hint is (name in READ_ONLY_TOOLS)
        assert annotations.destructive_hint is False
        assert annotations.idempotent_hint is (name in READ_ONLY_TOOLS)
        assert annotations.open_world_hint is (name != "whatsapp_draft_reply")


def test_descriptions_use_a_generic_owner_by_default(tmp_path: Path) -> None:
    tools = _tools(make_settings(tmp_path))
    assert "waiting on the user" in (tools["whatsapp_needs_reply"].description or "")
    assert "the user's latest message" in (tools["whatsapp_send_draft"].description or "")


def test_descriptions_use_the_configured_owner_name(tmp_path: Path) -> None:
    tools = _tools(make_settings(tmp_path, owner_name="Alex"))
    assert "waiting on Alex" in (tools["whatsapp_needs_reply"].description or "")
    assert "Show Alex the exact text" in (tools["whatsapp_draft_reply"].description or "")
    assert "Alex's latest message" in (tools["whatsapp_send_draft"].description or "")
