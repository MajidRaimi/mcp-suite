from __future__ import annotations

import pytest
from fakes.ax_backend import element
from fakes.trees import app_rig, window

from office_mcp.shared.ax import (
    ENHANCED_INTERFACE,
    MANUAL_ACCESSIBILITY,
    MENU_OPEN_CODES,
    AccessMode,
    custom_action_label,
)
from office_mcp.shared.errors import (
    AccessibilityError,
    AccessibilityModeError,
    AppNotRunningError,
    AxActionError,
    UiTimeoutError,
)


def test_require_trust_and_pid_raise_specific_errors() -> None:
    rig = app_rig(running=False)
    rig.backend.trusted = False

    with pytest.raises(AccessibilityError, match="Accessibility permission"):
        rig.app.require_trust()
    with pytest.raises(AppNotRunningError, match="Example App is not running"):
        rig.app.require_pid()
    assert rig.processes.launched == []


def test_web_mode_sets_manual_accessibility_once_per_pid() -> None:
    rig = app_rig(mode=AccessMode.WEB)

    rig.app.application()
    rig.app.application()

    assert rig.backend.sets_of(MANUAL_ACCESSIBILITY) == [True]
    assert rig.backend.sets_of(ENHANCED_INTERFACE) == []


def test_set_enhanced_tolerates_the_not_implemented_code_and_returns_previous() -> None:
    rig = app_rig(mode=AccessMode.WEB_ENHANCED)
    rig.root.attributes[ENHANCED_INTERFACE] = False
    rig.root.set_codes[ENHANCED_INTERFACE] = -25208
    rig.root.on_set[ENHANCED_INTERFACE] = lambda value: rig.root.attributes.update(
        {ENHANCED_INTERFACE: value}
    )

    previous = rig.app.set_enhanced(True)

    assert previous is False
    assert rig.root.attributes[ENHANCED_INTERFACE] is True


def test_set_enhanced_raises_when_the_value_does_not_stick() -> None:
    rig = app_rig(mode=AccessMode.WEB_ENHANCED)
    rig.root.on_set[ENHANCED_INTERFACE] = lambda value: None

    with pytest.raises(AccessibilityModeError, match="did not apply"):
        rig.app.set_enhanced(True)


def test_perform_tolerates_menu_codes_and_returns_them() -> None:
    target = element("AXButton", AXDescription="Move")
    target.action_codes["AXShowMenu"] = -25205
    rig = app_rig(window("Main", target))

    assert rig.app.perform(target, "AXShowMenu", tolerate=MENU_OPEN_CODES) == -25205
    with pytest.raises(AxActionError, match="AXShowMenu failed"):
        rig.app.perform(target, "AXShowMenu", "Move")


def test_custom_action_finds_the_raw_name() -> None:
    raw = "Name:Archive\nTarget:0x0\nSelector:(null)"
    row = element("AXRow", actions=("AXPress", raw, "Name:Settings...\nTarget:0x0"))
    rig = app_rig(window("Main", row))

    assert custom_action_label(raw) == "Archive"
    assert rig.app.custom_action(row, "Archive") == raw
    assert rig.app.custom_action(row, "Delete") is None


def test_window_or_launch_finds_existing_window_without_launching() -> None:
    main = window("Inbox")
    rig = app_rig(main)

    assert rig.app.window_or_launch(lambda w: rig.app.text(w, "AXTitle") == "Inbox") is main
    assert rig.processes.launched == []


def test_window_or_launch_launches_and_waits() -> None:
    rig = app_rig()
    opened = window("Inbox")

    def launch(_name: str) -> None:
        rig.root.attributes["AXWindows"] = [opened]

    rig.processes.on_launch = launch

    assert rig.app.window_or_launch(lambda w: True) is opened
    assert rig.processes.launched == ["Example App"]


def test_window_or_launch_times_out() -> None:
    rig = app_rig()

    with pytest.raises(UiTimeoutError, match="no matching window"):
        rig.app.window_or_launch(lambda w: True, timeout=2.0)
