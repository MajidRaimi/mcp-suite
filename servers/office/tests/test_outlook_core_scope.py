from __future__ import annotations

import pytest
from fakes.ax_backend import FakeBackend, FakeProcesses
from fakes.drivers import fake_runtime, outlook_rig
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.trees import outlook_rows
from fakes.workspace import FakeWorkspace

from office_mcp.features.outlook.ax import AxOutlook
from office_mcp.features.outlook.driver import build_outlook_driver
from office_mcp.features.outlook.rows import FolderLocator, RefRegistry
from office_mcp.features.outlook.window import FocusedTab
from office_mcp.shared.errors import OutlookUnavailableError
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.scope import ScopeOptions

ROW = "  Alice Example, Project Atlas budget,     9:15 AM,        Numbers attached"
OWNER_APP = 100


def refuse() -> None:
    raise OutlookUnavailableError("Outlook is not working right now.")


def test_health_failure_stops_before_any_ui_step() -> None:
    rig = outlook_rig(health=refuse)

    with pytest.raises(OutlookUnavailableError), rig.driver.scope(ScopeOptions(foreground=True)):
        pytest.fail("the scope body must not run")

    assert rig.gateway.calls == []
    assert rig.workspace.activations == []


def test_foreground_scope_activates_restores_view_then_gives_focus_back() -> None:
    gateway = FakeOutlookGateway(rows=outlook_rows(ROW), tab=FocusedTab.FOCUSED)
    rig = outlook_rig(gateway, workspace=FakeWorkspace(frontmost=OWNER_APP))
    pid = rig.app.app.require_pid()

    with rig.driver.scope(ScopeOptions(foreground=True)) as scope:
        assert rig.workspace.frontmost == pid
        gateway.switch_tab(FocusedTab.OTHER)
        gateway.select(gateway.rows[0])

    assert rig.health_calls == ["health"]
    assert gateway.tab is FocusedTab.FOCUSED
    assert gateway.selected == set()
    assert rig.workspace.frontmost == OWNER_APP
    assert scope.focus is FocusOutcome.RESTORED
    assert scope.view_notes == ("back to the focused tab", "cleared the selection")
    assert scope.cleanup.report == ("give the front back", "restore the Outlook view")


def test_read_scope_stays_in_the_background_until_asked() -> None:
    rig = outlook_rig(workspace=FakeWorkspace(frontmost=OWNER_APP))

    with rig.driver.scope(ScopeOptions(foreground=False, restore_view=False)) as scope:
        assert rig.workspace.activations == []
        scope.ensure_front()
        assert rig.workspace.frontmost == rig.app.app.require_pid()

    assert scope.focus is FocusOutcome.RESTORED
    assert rig.workspace.frontmost == OWNER_APP


def test_owner_moving_away_is_reported_and_not_undone() -> None:
    rig = outlook_rig(workspace=FakeWorkspace(frontmost=OWNER_APP))
    other_app = 300
    rig.workspace.running.add(other_app)

    with rig.driver.scope(ScopeOptions(foreground=True, restore_view=False)) as scope:
        rig.workspace.frontmost = other_app

    assert scope.focus is FocusOutcome.OWNER_MOVED
    assert rig.workspace.frontmost == other_app


def test_build_outlook_driver_wires_the_ax_gateway() -> None:
    runtime = fake_runtime(FakeBackend(), FakeProcesses())
    registry = RefRegistry()

    driver = build_outlook_driver(runtime, registry)

    assert isinstance(driver.gateway, AxOutlook)
    assert driver.registry is registry
    assert driver.app.bundle_id == "com.microsoft.Outlook"
    assert driver.focus.idle_seconds == runtime.settings.idle_seconds


class FrontWatchingGateway(FakeOutlookGateway):
    def __init__(self, desk: FakeWorkspace) -> None:
        super().__init__(rows=outlook_rows(ROW), tab=FocusedTab.FOCUSED)
        self.desk = desk
        self.front_at_switch: list[int | None] = []

    def switch_tab(self, tab: FocusedTab) -> None:
        self.front_at_switch.append(self.desk.frontmost)
        super().switch_tab(tab)


def test_a_background_read_brings_outlook_forward_before_switching_tabs() -> None:
    desk = FakeWorkspace(frontmost=OWNER_APP)
    gateway = FrontWatchingGateway(desk)
    rig = outlook_rig(gateway, workspace=desk)
    pid = rig.app.app.require_pid()

    with rig.driver.scope(ScopeOptions(foreground=False)) as scope:
        rig.driver.rows.rows(FolderLocator("", FocusedTab.OTHER))

    assert gateway.front_at_switch == [pid, pid]
    assert gateway.tab is FocusedTab.FOCUSED
    assert rig.workspace.frontmost == OWNER_APP
    assert scope.focus is FocusOutcome.RESTORED


def test_rows_outside_a_scope_switch_tabs_without_activating() -> None:
    gateway = FakeOutlookGateway(rows=outlook_rows(ROW), tab=FocusedTab.FOCUSED)
    rig = outlook_rig(gateway, workspace=FakeWorkspace(frontmost=OWNER_APP))

    rig.driver.rows.rows(FolderLocator("", FocusedTab.OTHER))

    assert rig.workspace.activations == []
    assert gateway.tab is FocusedTab.OTHER
