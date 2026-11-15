from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

from office_mcp.cli import ALIASES, BOOTSTRAP, build_parser, command_name, main
from office_mcp.container import all_specs, build_container
from office_mcp.shared.settings import Settings


@pytest.fixture(autouse=True)
def clean_environment(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    for name in list(os.environ):
        if name.startswith("OFFICE_MCP_"):
            monkeypatch.delenv(name)
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("OFFICE_MCP_TIMEZONE", "UTC")
    monkeypatch.setenv("OFFICE_MCP_LOCK_FILE", str(tmp_path / "ui.lock"))
    return tmp_path


def every_spec_parser(tmp_path: Path) -> tuple[list[str], list[str]]:
    settings = Settings.from_env(
        {
            "OFFICE_MCP_TIMEZONE": "UTC",
            "OFFICE_MCP_TOOLS": "all",
            "OFFICE_MCP_EXPERIMENTAL": "all",
            "OFFICE_MCP_DOWNLOAD_DIR": str(tmp_path / "downloads"),
        }
    )
    specs = all_specs(build_container(settings))
    return [spec.name for spec in specs], [command_name(spec) for spec in specs]


def parsed(*argv: str) -> dict[str, object]:
    settings = Settings.from_env(
        {"OFFICE_MCP_TIMEZONE": "UTC", "OFFICE_MCP_DOWNLOAD_DIR": "/tmp/office-downloads"}
    )
    return vars(build_parser(all_specs(build_container(settings))).parse_args(list(argv)))


def test_parser_has_a_command_for_every_spec(tmp_path: Path) -> None:
    names, commands = every_spec_parser(tmp_path)

    assert len(names) == 35
    for name, command in zip(names, commands, strict=True):
        assert parsed(command, *required_arguments(name))["tool"] == name


def required_arguments(name: str) -> list[str]:
    samples = {
        "office_check_text": ["hello there", "teams"],
        "outlook_read": ["0123456789ab"],
        "outlook_search": ["budget"],
        "outlook_people": ["Jane"],
        "outlook_save_attachment": ["0123456789ab", "plan.pdf"],
        "outlook_archive": ["--refs", "0123456789ab"],
        "outlook_delete": ["--refs", "0123456789ab"],
        "outlook_move": ["Archive", "--refs", "0123456789ab"],
        "outlook_mark": ["--refs", "0123456789ab", "--read"],
        "outlook_flag": ["--refs", "0123456789ab"],
        "outlook_send": ["Project Atlas budget", "Hi", "--to", "jane@example.com"],
        "outlook_reply": ["0123456789ab", "Thanks"],
        "outlook_forward": ["0123456789ab", "--to", "jane@example.com"],
        "calendar_event": ["0123456789ab"],
        "calendar_respond": ["0123456789ab", "accept"],
        "teams_read": ["Jane Doe"],
        "teams_search": ["atlas"],
        "teams_channel_read": ["Example Team", "General"],
        "teams_send": ["48:notes", "hello"],
        "teams_reply": ["48:notes", "1700000000000", "hello"],
        "teams_react": ["48:notes", "1700000000000", "like"],
        "teams_delete_message": ["48:notes", "1700000000000"],
        "teams_mark_unread": ["48:notes"],
    }
    return samples.get(name, [])


def test_old_command_names_still_work() -> None:
    for name, alias in ALIASES.items():
        assert parsed(alias, *required_arguments(name))["tool"] == name


def test_options_follow_the_parameter_types() -> None:
    send = parsed(
        "outlook-send",
        "Project Atlas budget",
        "Numbers attached",
        "--to",
        "jane@example.com",
        "--to",
        "omar@example.com",
        "--cc",
        "alex@example.com",
        "--dry-run",
    )
    conversations = parsed("mail", "--folder", "Archive", "--limit", "5", "--view", "both")

    assert send["to"] == ["jane@example.com", "omar@example.com"]
    assert send["cc"] == ["alex@example.com"]
    assert send["bcc"] is None
    assert send["dry_run"] is True
    assert conversations["limit"] == 5
    assert conversations["view"] == "both"
    assert conversations["days"] is None
    assert parsed("read-mail", "0123456789ab", "--mark-read")["keep_unread"] is False
    assert parsed("read-mail", "0123456789ab")["keep_unread"] is True
    assert parsed("outlook-mark", "--refs", "0123456789ab", "--no-read")["read"] is False


def test_literal_parameters_refuse_other_values() -> None:
    with pytest.raises(SystemExit):
        parsed("outlook-conversations", "--view", "sideways")
    with pytest.raises(SystemExit):
        parsed("teams-react", "48:notes", "1700000000000", "thumbs")


def test_bootstrap_command_has_the_documented_defaults() -> None:
    arguments = parsed(BOOTSTRAP)

    assert arguments["emails"] == 40
    assert arguments["chats"] == 15
    assert arguments["per_chat"] == 20
    assert arguments["out"] == Path("~/.config/office-mcp")
    assert arguments["keep_corpus"] is False


def test_disabled_commands_are_refused(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as stopped:
        main(["outlook-archive", "--refs", "0123456789ab"])

    assert stopped.value.code == 1
    assert "outlook_archive is turned off" in json.loads(capsys.readouterr().out)["error"]


def test_enabled_commands_run_the_tool(capsys: pytest.CaptureFixture[str]) -> None:
    main(["office-check-text", "Short update on the plan", "teams"])

    record = json.loads(capsys.readouterr().out)
    assert record["ok"] is True
    assert record["final_text"] == "Short update on the plan"


def test_unknown_tool_names_stop_the_cli(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    monkeypatch.setenv("OFFICE_MCP_DISABLED_TOOLS", "outlook_nope")

    with pytest.raises(SystemExit):
        main(["status"])

    assert "unknown tools" in json.loads(capsys.readouterr().out)["error"]


def test_bootstrap_refuses_an_out_dir_outside_the_config_folder(
    clean_environment: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    with pytest.raises(SystemExit):
        main([BOOTSTRAP, "--out", str(clean_environment / "elsewhere")])

    assert "only inside" in json.loads(capsys.readouterr().out)["error"]
