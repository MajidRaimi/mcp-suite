from __future__ import annotations

import pytest
from fakes.clock import FakeClock
from fakes.drivers import fake_runtime, outlook_rig
from fakes.lock import MemoryUiLock
from fakes.p2_outlook import ActionsRig, OrganizeOutlook, actions_rig, atlas
from fakes.workspace import FakeWorkspace
from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.features.outlook.actions.factory import build_actions
from office_mcp.features.outlook.actions.tools import ActionsTools
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
    annotations_for,
)
from office_mcp.shared.runner import Deadline, Lazy, Runner

OWNER = Owner("Alex Example")
OWNER_APP = 100


def tools_for(rig: ActionsRig, lock: MemoryUiLock | None = None) -> ActionsTools:
    runner = Runner(lock or MemoryUiLock(), FakeClock(), lock_timeout=20.0, stale_after=10.0)
    return ActionsTools(runner, Lazy(lambda: rig.service), OWNER)


def test_specs_cover_tools_twelve_to_sixteen() -> None:
    specs = {spec.name: spec for spec in tools_for(actions_rig({})).specs()}

    assert list(specs) == [
        "outlook_archive",
        "outlook_delete",
        "outlook_move",
        "outlook_mark",
        "outlook_flag",
    ]
    assert {spec.group for spec in specs.values()} == {ToolGroup.ORGANIZE}
    assert specs["outlook_delete"].effect is Effect.SOFT_DELETES
    assert annotations_for(specs["outlook_delete"].effect).destructive_hint is True
    assert specs["outlook_archive"].effect is Effect.ORGANIZES


def test_every_organize_tool_is_stable_after_the_live_run() -> None:
    specs = tools_for(actions_rig({})).specs()

    for spec in specs:
        assert spec.stability is Stability.STABLE
        assert not spec.description.startswith(EXPERIMENTAL_PREFIX)


def test_descriptions_name_the_owner_and_carry_no_dashes() -> None:
    for spec in tools_for(actions_rig({})).specs():
        assert "Alex Example's mailbox" in spec.description
        assert chr(0x2013) not in spec.description
        assert chr(0x2014) not in spec.description
        assert "explicit request" in spec.description


def test_flag_tool_runs_through_the_runner_and_returns_the_record() -> None:
    rig = actions_rig({"m1": atlas()})
    lock = MemoryUiLock()

    record = tools_for(rig, lock).outlook_flag([rig.key("m1")], "tomorrow")

    assert record["action"] == "flag"
    assert record["succeeded"] == 1
    assert record["focus_restored"] == "restored"
    assert (lock.acquired, lock.released) == (1, 1)


def test_bad_refs_become_tool_errors() -> None:
    rig = actions_rig({"m1": atlas()})

    with pytest.raises(ToolError, match="use the key"):
        tools_for(rig).outlook_delete(["2"])

    assert rig.outlook.calls == []


def test_factory_wires_the_driver_and_its_foreground_scope() -> None:
    outlook = OrganizeOutlook({"m1": atlas()})
    outlook.actions = {"m1": ("Archive",)}
    rig = outlook_rig(outlook, workspace=FakeWorkspace(frontmost=OWNER_APP))
    service = build_actions(fake_runtime(), rig.driver)

    result = service.archive([outlook.key_of("m1")], None, Deadline.after(FakeClock(), 60.0))

    assert result.items[0].verified is True
    assert result.focus is FocusOutcome.RESTORED
    assert rig.workspace.frontmost == OWNER_APP
    assert rig.health_calls == ["health"]
