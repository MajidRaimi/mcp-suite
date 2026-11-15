from __future__ import annotations

import pytest
from fakes.clock import fake_waiter
from fakes.trees import TEST_PID, AppRig, app_rig, window
from fakes.workspace import FakeWorkspace

from office_mcp.shared.errors import ActivationError, BusyError, SessionUnavailableError
from office_mcp.shared.focus import FocusOutcome, FocusPolicy, ForegroundSession, FrontHook

OWNER_APP = 100
POLICY = FocusPolicy(restore=True, idle_seconds=1.5)


def session(
    rig: AppRig, workspace: FakeWorkspace, policy: FocusPolicy = POLICY
) -> ForegroundSession:
    return ForegroundSession(workspace, rig.app, fake_waiter(rig.clock), policy)


def setup() -> tuple[AppRig, FakeWorkspace]:
    rig = app_rig(window("Main", AXMain=True))
    workspace = FakeWorkspace(frontmost=OWNER_APP, running={OWNER_APP, TEST_PID})
    return rig, workspace


def test_activates_then_restores_the_previous_app() -> None:
    rig, workspace = setup()

    with session(rig, workspace) as foreground:
        assert workspace.frontmost == TEST_PID
        foreground.ensure_front()

    assert workspace.frontmost == OWNER_APP
    assert foreground.outcome is FocusOutcome.RESTORED
    assert rig.backend.sets_of("AXFrontmost") == [True]
    assert len(rig.backend.actions_of("AXRaise")) == 1


def test_reports_owner_moved_when_the_owner_switched_apps() -> None:
    rig, workspace = setup()

    with session(rig, workspace) as foreground:
        workspace.frontmost = 300

    assert foreground.outcome is FocusOutcome.OWNER_MOVED
    assert workspace.frontmost == 300


def test_not_needed_when_already_front_and_kept_when_restore_disabled() -> None:
    rig, workspace = setup()
    workspace.frontmost = TEST_PID

    with session(rig, workspace) as already:
        pass
    assert already.outcome is FocusOutcome.NOT_NEEDED

    workspace.frontmost = OWNER_APP
    with session(rig, workspace, FocusPolicy(restore=False, idle_seconds=0)) as kept:
        pass
    assert kept.outcome is FocusOutcome.KEPT
    assert workspace.frontmost == TEST_PID


def test_retries_when_focus_is_stolen() -> None:
    rig, workspace = setup()
    workspace.refuse_activations = 2

    with session(rig, workspace):
        assert workspace.frontmost == TEST_PID

    assert workspace.activations[:3] == [TEST_PID, TEST_PID, TEST_PID]


def test_raises_activation_error_after_three_attempts() -> None:
    rig, workspace = setup()
    workspace.refuse_activations = 10

    with pytest.raises(ActivationError, match="Example App"), session(rig, workspace):
        pass
    assert workspace.activations == [TEST_PID, TEST_PID, TEST_PID]


def test_ensure_front_reactivates_after_a_steal() -> None:
    rig, workspace = setup()

    with session(rig, workspace) as foreground:
        workspace.frontmost = 555
        foreground.ensure_front()
        assert workspace.frontmost == TEST_PID


@pytest.mark.parametrize(("locked", "secure"), [(True, False), (False, True)])
def test_refuses_locked_screen_and_secure_input(locked: bool, secure: bool) -> None:
    rig, workspace = setup()
    workspace.locked = locked
    workspace.secure = secure

    with pytest.raises(SessionUnavailableError), session(rig, workspace):
        pass
    assert workspace.activations == []


def test_waits_for_idle_then_proceeds() -> None:
    rig, workspace = setup()
    workspace.idle = lambda: rig.clock.now

    with session(rig, workspace):
        pass

    assert rig.clock.now >= 1.5
    assert workspace.activations[0] == TEST_PID


def test_busy_when_the_owner_keeps_typing() -> None:
    rig, workspace = setup()
    workspace.idle = lambda: 0.1

    with pytest.raises(BusyError, match="owner is using the Mac"), session(rig, workspace):
        pass
    assert rig.clock.now == pytest.approx(10.0)
    assert workspace.activations == []


def test_front_hook_calls_the_latest_binding_until_it_is_released() -> None:
    hook = FrontHook()
    calls: list[str] = []
    hook.bring_front()

    release_outer = hook.bind(lambda: calls.append("outer"))
    release_inner = hook.bind(lambda: calls.append("inner"))
    hook.bring_front()
    release_inner()
    hook.bring_front()
    release_outer()
    hook.bring_front()

    assert calls == ["inner", "outer"]
