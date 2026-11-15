from __future__ import annotations

import pytest
from fakes.ax_backend import element
from fakes.trees import app_rig, button, window

from office_mcp.shared.errors import ElementNotFoundError, UiTimeoutError
from office_mcp.shared.tree import by_role_and_label, find, find_all
from office_mcp.shared.windows import WindowLocator, WindowProbe


def test_find_is_breadth_first_with_depth_cap_and_prune() -> None:
    shallow = button("Send")
    deep = element("AXGroup", element("AXGroup", button("Deep")))
    table = element("AXTable", button("Inside table"))
    root = window("Main", deep, table, shallow)
    rig = app_rig(root)
    is_button = by_role_and_label(rig.app, "AXButton", "Send")

    assert find(rig.app, root, is_button) is shallow
    assert find(rig.app, root, by_role_and_label(rig.app, "AXButton", "Deep"), max_depth=2) is None
    pruned = find(
        rig.app,
        root,
        by_role_and_label(rig.app, "AXButton", "Inside table"),
        prune=lambda e: rig.app.text(e, "AXRole") == "AXTable",
    )
    assert pruned is None


def test_find_skips_a_self_child() -> None:
    field = element("AXTextField", AXDescription="Search")
    field.extra_children.append(field)
    root = window("Main", field)
    rig = app_rig(root)

    found = find_all(rig.app, root, lambda e: rig.app.text(e, "AXRole") == "AXTextField")

    assert found == [field]


def test_find_all_respects_limit() -> None:
    root = window("Main", *(button(f"b{n}") for n in range(5)))
    rig = app_rig(root)

    found = find_all(rig.app, root, lambda e: rig.app.text(e, "AXRole") == "AXButton", limit=3)

    assert len(found) == 3


def kind(probe: WindowProbe) -> str:
    return "compose" if "Send" in probe.landmarks else "main"


def test_window_locator_probes_landmarks_and_flags() -> None:
    main = window(
        "Inbox",
        element("AXGroup", button("Mail", AXHelp="Mail")),
        element("AXTable", button("Hidden row"), AXDescription="Message List"),
        AXMain=True,
        AXSubrole="AXStandardWindow",
    )
    rig = app_rig(main)

    probe = WindowLocator(rig.app).probes()[0]

    assert probe.main and not probe.focused
    assert probe.subrole == "AXStandardWindow"
    assert {"Mail", "Message List"} <= probe.landmarks
    assert "Hidden row" not in probe.landmarks


def test_window_locator_require_and_wait_new() -> None:
    main = window("Inbox", button("Mail"))
    rig = app_rig(main)
    locator = WindowLocator(rig.app)
    before = locator.probes()

    with pytest.raises(ElementNotFoundError, match="no compose"):
        locator.require(kind, "compose", "no compose window")
    with pytest.raises(UiTimeoutError):
        locator.wait_new(before, kind, "compose")

    compose = window("Untitled", button("Send"))
    rig.root.attributes["AXWindows"] = [main, compose]

    assert locator.wait_new(before, kind, "compose").handle is compose
    assert locator.require(kind, "main", "no main").handle is main
