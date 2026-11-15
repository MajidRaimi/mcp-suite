from __future__ import annotations

import pytest
from fakes.ax_backend import FakeBackend, FakeProcesses, element
from fakes.drivers import fake_runtime, teams_rig
from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import ACTIVITY_APP_ID, TEST_PID, chat_row, message, teams_window
from fakes.workspace import FakeWorkspace

from office_mcp.features.teams.ax import AxTeams
from office_mcp.features.teams.driver import build_teams_driver
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.ax import ENHANCED_INTERFACE
from office_mcp.shared.errors import CleanupError
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.node import Node
from office_mcp.shared.scope import ScopeOptions

BOB = "19:bob@unq.gbl.spaces"
ALICE = "19:alice@unq.gbl.spaces"
ROWS = [
    chat_row(BOB, "Bob Sample", "ok", "10/19"),
    chat_row(ALICE, "Alice Example", "hi", "4:09 AM"),
]
OWNER_APP = 100
QUIET = ScopeOptions(foreground=False, restore_view=False)


def chat_open(thread: str, name: str, rows: list[Node] = ROWS) -> Node:
    body = message("1790000000000", name, "hello", "October 19, 2026 9:00 AM")
    return teams_window(f"Chat | {name}", rows, body, thread=thread)


def activity() -> Node:
    return teams_window("Activity", [], [], selected_app=ACTIVITY_APP_ID)


def enhanced_values(rig_backend: FakeBackend) -> list[object]:
    return rig_backend.sets_of(ENHANCED_INTERFACE)


def test_enhanced_mode_is_switched_on_then_back_off() -> None:
    rig = teams_rig(FakeTeamsGateway([chat_open(BOB, "Bob Sample")]))

    with rig.driver.scope(QUIET) as scope:
        assert rig.app.root.attributes[ENHANCED_INTERFACE] is True

    assert enhanced_values(rig.app.backend) == [True, False]
    assert scope.cleanup.report == ("close Teams menus", "restore the Teams accessibility mode")
    assert rig.gateway.close_calls == 2


def test_enhanced_mode_that_was_already_on_stays_on() -> None:
    rig = teams_rig(FakeTeamsGateway([chat_open(BOB, "Bob Sample")]))
    rig.app.root.attributes[ENHANCED_INTERFACE] = True

    with rig.driver.scope(QUIET):
        pass

    assert rig.app.root.attributes[ENHANCED_INTERFACE] is True


def test_view_restore_returns_to_the_rail_app_and_the_open_thread() -> None:
    gateway = FakeTeamsGateway([chat_open(BOB, "Bob Sample")])
    rig = teams_rig(gateway)

    def press(node: Node) -> None:
        if node.handle == "rail:Activity":
            gateway.show(activity())
        elif node.handle == "rail:Chat":
            gateway.show(chat_open(ALICE, "Alice Example"))
        elif node.handle == BOB:
            gateway.show(chat_open(BOB, "Bob Sample"))

    gateway.on_press = press
    with rig.driver.scope(ScopeOptions(foreground=False)) as scope:
        assert scope.show_app(TeamsApp.ACTIVITY).title == "Activity"

    assert gateway.pressed == ["rail:Activity", "rail:Chat", BOB]
    assert gateway.snapshot().title == "Chat | Bob Sample"


def test_view_restore_waits_for_the_chat_pane_before_looking_in_the_list() -> None:
    gateway = FakeTeamsGateway([chat_open(BOB, "Bob Sample", rows=ROWS[1:])])
    rig = teams_rig(gateway)
    loading = teams_window("Chat", ROWS[1:], [])

    def press(node: Node) -> None:
        if node.handle == "rail:Activity":
            gateway.show(activity())
        elif node.handle == "rail:Chat":
            gateway.show(loading, loading, chat_open(BOB, "Bob Sample", rows=ROWS[1:]))

    gateway.on_press = press
    with rig.driver.scope(ScopeOptions(foreground=False)) as scope:
        scope.show_app(TeamsApp.ACTIVITY)

    assert gateway.pressed == ["rail:Activity", "rail:Chat"]
    assert gateway.snapshot().title == "Chat | Bob Sample"


def test_view_restore_fails_loudly_when_the_old_chat_is_gone() -> None:
    gateway = FakeTeamsGateway([chat_open(BOB, "Bob Sample")])
    rig = teams_rig(gateway)

    with (
        pytest.raises(CleanupError, match="no longer in the list"),
        rig.driver.scope(ScopeOptions(foreground=False)),
    ):
        gateway.show(chat_open(ALICE, "Alice Example", rows=ROWS[1:]))

    assert rig.app.root.attributes[ENHANCED_INTERFACE] is False


def test_foreground_scope_raises_teams_and_gives_the_front_back() -> None:
    rig = teams_rig(
        FakeTeamsGateway([chat_open(BOB, "Bob Sample")]),
        workspace=FakeWorkspace(frontmost=OWNER_APP),
    )

    with rig.driver.scope(ScopeOptions(foreground=True, restore_view=False)) as scope:
        assert rig.workspace.frontmost == TEST_PID

    assert rig.workspace.frontmost == OWNER_APP
    assert scope.focus is FocusOutcome.RESTORED


def test_quiet_scope_raises_teams_only_when_asked() -> None:
    rig = teams_rig(
        FakeTeamsGateway([chat_open(BOB, "Bob Sample")]),
        workspace=FakeWorkspace(frontmost=OWNER_APP),
    )

    with rig.driver.scope(QUIET) as scope:
        assert scope.focus is FocusOutcome.NOT_NEEDED
        assert rig.workspace.activations == []
        scope.ensure_front()
        assert rig.workspace.frontmost == TEST_PID

    assert rig.workspace.frontmost == OWNER_APP


def test_build_teams_driver_resets_a_stale_enhanced_mode() -> None:
    backend = FakeBackend()
    root = backend.register(TEST_PID, element("AXApplication"))
    root.attributes[ENHANCED_INTERFACE] = True
    runtime = fake_runtime(backend, FakeProcesses({"com.microsoft.teams2": TEST_PID}))

    driver = build_teams_driver(runtime)

    assert isinstance(driver.gateway, AxTeams)
    assert root.attributes[ENHANCED_INTERFACE] is False
    assert driver.zone == runtime.settings.zone.info
