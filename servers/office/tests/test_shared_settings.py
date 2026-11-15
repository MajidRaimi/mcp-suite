from __future__ import annotations

from pathlib import Path

import pytest

from office_mcp.shared.errors import SettingsError
from office_mcp.shared.registry import ToolGroup
from office_mcp.shared.settings import Settings, parse_bool, parse_groups


def test_new_defaults() -> None:
    settings = Settings.from_env({})

    assert settings.tool_policy.groups == frozenset({ToolGroup.READ})
    assert settings.tool_policy.disabled == frozenset()
    assert settings.tool_policy.experimental == frozenset()
    assert not settings.tool_policy.all_experimental
    assert settings.persona_file == Path.home() / ".config" / "office-mcp" / "persona.md"
    assert not settings.persona_explicit and not settings.persona_required
    assert settings.send_only_to is None
    assert settings.self_chat_thread == "48:notes"
    assert settings.restore_frontmost
    assert (settings.action_timeout, settings.idle_seconds) == (12.0, 1.5)
    assert settings.lock_file == Path.home() / "Library" / "Caches" / "office-mcp" / "ui.lock"
    assert (settings.lock_timeout, settings.stale_after) == (20.0, 10.0)
    assert settings.download_dir is None
    assert settings.teams_max_members == 25
    assert settings.owner_email is None


def test_reads_every_new_variable(tmp_path: Path) -> None:
    settings = Settings.from_env(
        {
            "OFFICE_MCP_TOOLS": "organize, SEND",
            "OFFICE_MCP_DISABLED_TOOLS": "calendar_respond, ",
            "OFFICE_MCP_EXPERIMENTAL": "outlook_move,teams_mark_unread",
            "OFFICE_MCP_PERSONA_FILE": str(tmp_path / "persona.md"),
            "OFFICE_MCP_PERSONA_REQUIRED": "yes",
            "OFFICE_MCP_SEND_ONLY_TO": "jane@example.com, teams:48:notes",
            "OFFICE_MCP_SELF_CHAT_THREAD": " 19:self@unq.gbl.spaces ",
            "OFFICE_MCP_RESTORE_FRONTMOST": "off",
            "OFFICE_MCP_ACTION_TIMEOUT": "30",
            "OFFICE_MCP_IDLE_SECONDS": "0",
            "OFFICE_MCP_LOCK_FILE": "~/locks/ui.lock",
            "OFFICE_MCP_LOCK_TIMEOUT": "45.5",
            "OFFICE_MCP_STALE_AFTER": "5",
            "OFFICE_MCP_DOWNLOAD_DIR": str(tmp_path / "downloads"),
            "OFFICE_MCP_TEAMS_MAX_MEMBERS": "8",
            "OFFICE_MCP_OWNER_EMAIL": "Jane@Example.com",
        }
    )

    policy = settings.tool_policy
    assert policy.groups == frozenset(ToolGroup)
    assert policy.disabled == frozenset({"calendar_respond"})
    assert policy.experimental == frozenset({"outlook_move", "teams_mark_unread"})
    assert settings.persona_file == tmp_path / "persona.md"
    assert settings.persona_explicit and settings.persona_required
    assert settings.send_only_to is not None
    assert settings.send_only_to.entries() == ("jane@example.com", "teams:48:notes")
    assert settings.self_chat_thread == "19:self@unq.gbl.spaces"
    assert not settings.restore_frontmost
    assert (settings.action_timeout, settings.idle_seconds) == (30.0, 0.0)
    assert settings.lock_file == Path.home() / "locks" / "ui.lock"
    assert (settings.lock_timeout, settings.stale_after) == (45.5, 5.0)
    assert settings.download_dir == tmp_path / "downloads"
    assert settings.teams_max_members == 8
    assert settings.owner_email == "jane@example.com"


def test_tools_all_and_experimental_all() -> None:
    policy = Settings.from_env(
        {"OFFICE_MCP_TOOLS": "ALL", "OFFICE_MCP_EXPERIMENTAL": "all"}
    ).tool_policy

    assert policy.groups == frozenset(ToolGroup)
    assert policy.all_experimental
    assert policy.experimental == frozenset()


def test_read_only_wins() -> None:
    policy = Settings.from_env(
        {
            "OFFICE_MCP_READ_ONLY": "1",
            "OFFICE_MCP_TOOLS": "all",
            "OFFICE_MCP_EXPERIMENTAL": "outlook_move",
            "OFFICE_MCP_DISABLED_TOOLS": "teams_read",
        }
    ).tool_policy

    assert policy.groups == frozenset({ToolGroup.READ})
    assert policy.experimental == frozenset()
    assert not policy.all_experimental
    assert policy.disabled == frozenset({"teams_read"})


def test_parse_groups_always_adds_read() -> None:
    assert parse_groups(None) == frozenset({ToolGroup.READ})
    assert parse_groups("send") == frozenset({ToolGroup.READ, ToolGroup.SEND})


@pytest.mark.parametrize("raw", ["1", "TRUE", "yes", "On"])
def test_parse_bool_true_words(raw: str) -> None:
    assert parse_bool(raw, "X", False) is True


@pytest.mark.parametrize("raw", ["0", "false", "NO", "off"])
def test_parse_bool_false_words(raw: str) -> None:
    assert parse_bool(raw, "X", True) is False


@pytest.mark.parametrize(
    ("variable", "value", "message"),
    [
        ("OFFICE_MCP_TOOLS", "read,admin", "not one of"),
        ("OFFICE_MCP_READ_ONLY", "maybe", "true or false"),
        ("OFFICE_MCP_PERSONA_REQUIRED", "sure", "true or false"),
        ("OFFICE_MCP_RESTORE_FRONTMOST", "2", "true or false"),
        ("OFFICE_MCP_SEND_ONLY_TO", "not-an-address", "must be user@domain"),
        ("OFFICE_MCP_SEND_ONLY_TO", " , ", "lists no entries"),
        ("OFFICE_MCP_ACTION_TIMEOUT", "0.5", "between 1 and 120"),
        ("OFFICE_MCP_ACTION_TIMEOUT", "soon", "number of seconds"),
        ("OFFICE_MCP_IDLE_SECONDS", "-1", "between 0 and 30"),
        ("OFFICE_MCP_IDLE_SECONDS", "nan", "between 0 and 30"),
        ("OFFICE_MCP_LOCK_TIMEOUT", "301", "between 1 and 300"),
        ("OFFICE_MCP_STALE_AFTER", "0", "between 1 and 300"),
        ("OFFICE_MCP_DOWNLOAD_DIR", "relative/downloads", "absolute path"),
        ("OFFICE_MCP_TEAMS_MAX_MEMBERS", "1", "between 2 and 1000"),
        ("OFFICE_MCP_TEAMS_MAX_MEMBERS", "many", "whole number"),
        ("OFFICE_MCP_OWNER_EMAIL", "jane at example.com", "email address"),
    ],
)
def test_invalid_new_values_raise(variable: str, value: str, message: str) -> None:
    with pytest.raises(SettingsError, match=message):
        Settings.from_env({variable: value})
