from __future__ import annotations

import inspect

import pytest
from fakes.clock import FakeClock
from fakes.lock import MemoryUiLock
from fakes.p3_compose import compose_rig
from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.features.outlook.compose.tools import ComposeTools
from office_mcp.shared.drafts import Channel
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
    annotations_for,
)
from office_mcp.shared.runner import SEND_BUDGET, Lazy, Runner

RULES = "Persona rules for this text: no dashes."


def rules_text(channel: Channel) -> str:
    return f"{RULES} ({channel.value})"


def tools_for_rig() -> tuple[ComposeTools, object]:
    rig = compose_rig()
    clock = FakeClock()
    runner = Runner(MemoryUiLock(), clock, lock_timeout=20, stale_after=10)
    tools = ComposeTools(runner, Lazy(lambda: rig.service), Owner("Jane Doe"), rules_text)
    return tools, rig


def test_specs_cover_tools_17_to_20() -> None:
    tools, _ = tools_for_rig()

    specs = {spec.name: spec for spec in tools.specs()}

    assert list(specs) == ["outlook_send", "outlook_reply", "outlook_forward", "outlook_undo_send"]
    for spec in specs.values():
        assert spec.group is ToolGroup.SEND
        assert spec.effect is Effect.SENDS
        assert "\u2013" not in spec.description
        assert "\u2014" not in spec.description
    assert annotations_for(Effect.SENDS).open_world_hint is True
    assert specs["outlook_undo_send"].stability is Stability.EXPERIMENTAL
    assert specs["outlook_undo_send"].description.startswith(EXPERIMENTAL_PREFIX)
    for name in ("outlook_send", "outlook_reply", "outlook_forward"):
        assert specs[name].stability is Stability.STABLE
        assert f"{RULES} (email)" in specs[name].description
        assert specs[name].description.endswith("Call office_check_text first when unsure.")
        assert "Jane Doe" in specs[name].description


def test_parameters_match_the_catalogue() -> None:
    tools, _ = tools_for_rig()

    send = inspect.signature(tools.outlook_send).parameters
    reply = inspect.signature(tools.outlook_reply).parameters
    forward = inspect.signature(tools.outlook_forward).parameters

    assert list(send) == [
        "to",
        "subject",
        "body",
        "cc",
        "bcc",
        "allow_language_mismatch",
        "allow_duplicate",
        "dry_run",
        "sensitivity",
    ]
    assert list(reply) == [
        "ref",
        "body",
        "reply_all",
        "cc",
        "bcc",
        "folder",
        "allow_language_mismatch",
        "allow_duplicate",
        "dry_run",
        "sensitivity",
    ]
    assert list(forward) == [
        "ref",
        "to",
        "body",
        "cc",
        "bcc",
        "folder",
        "allow_duplicate",
        "dry_run",
        "sensitivity",
    ]
    assert send["dry_run"].default is False
    assert send["sensitivity"].default is None
    assert forward["body"].default == ""


def test_dry_run_send_returns_a_record() -> None:
    tools, _ = tools_for_rig()

    record = tools.outlook_send(
        to=["bob@example.com"], subject="Project Atlas budget", body="Hello", dry_run=True
    )

    assert record["status"] == "dry_run"
    assert record["text_sent"] == "Hello"
    assert record["persona"]["version"] == "test"


def test_dry_run_send_reports_the_sensitivity_path() -> None:
    tools, _ = tools_for_rig()

    record = tools.outlook_send(
        to=["bob@example.com"],
        subject="Project Atlas budget",
        body="Hello",
        dry_run=True,
        sensitivity=["Confidential", "Internal"],
    )

    assert record["sensitivity"] == ["Confidential", "Internal"]


def test_errors_become_tool_errors() -> None:
    tools, _ = tools_for_rig()

    with pytest.raises(ToolError, match="outlook_people"):
        tools.outlook_send(to=["Bob Sample"], subject="Hi", body="Hello", dry_run=True)


def test_undo_without_a_send_is_a_tool_error() -> None:
    tools, _ = tools_for_rig()

    with pytest.raises(ToolError, match="nothing to undo"):
        tools.outlook_undo_send()


def test_send_budget_covers_a_reply_in_a_large_mailbox() -> None:
    assert SEND_BUDGET >= 90.0
