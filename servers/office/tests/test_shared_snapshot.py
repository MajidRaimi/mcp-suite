from __future__ import annotations

from fakes.ax_backend import FakeElement, element
from fakes.trees import app_rig, window

from office_mcp.shared.snapshot import Snapshotter


def wide_tree(count: int) -> FakeElement:
    return window(
        "Main",
        *(element("AXGroup", element("AXStaticText", AXValue=f"item {n}")) for n in range(count)),
    )


def test_snapshot_keeps_every_node_when_all_elements_hash_equal() -> None:
    root = wide_tree(50)
    rig = app_rig(root)

    snapshot = rig.app.snapshot(root)

    assert len({hash(kid) for kid in root.children}) == 1
    assert sum(1 for _ in snapshot.walk()) == 101
    assert "item 49" in snapshot.text()


def test_child_equal_to_its_parent_is_skipped() -> None:
    field = element("AXTextField", element("AXButton", AXDescription="Perform Search"))
    field.extra_children.append(field)
    root = window("Main", field)
    rig = app_rig(root)

    snapshot = rig.app.snapshot(root)

    roles = [node.role for node in snapshot.walk()]
    assert roles == ["AXWindow", "AXTextField", "AXButton"]


def test_depth_cap_stops_descent() -> None:
    leaf = element("AXStaticText", AXValue="deep")
    node = leaf
    for _ in range(10):
        node = element("AXGroup", node)
    rig = app_rig(node)

    shallow = rig.app.snapshot(node, max_depth=3)

    assert sum(1 for _ in shallow.walk()) == 4
    assert shallow.text() == ""


def test_node_budget_caps_the_walk() -> None:
    root = wide_tree(30)
    rig = app_rig(root)

    snapshot = Snapshotter(rig.app, budget=10).build(root)

    assert sum(1 for _ in snapshot.walk()) == 10


def test_snapshot_reads_new_attributes() -> None:
    button = element(
        "AXButton",
        AXSubrole="AXToggle",
        AXHelp="Calendar",
        AXEnabled=False,
        AXDisclosureLevel=2,
        AXSelected=True,
    )
    rig = app_rig(window("Main", button))

    node = rig.app.snapshot(button)

    assert (node.subrole, node.help, node.enabled, node.level, node.selected) == (
        "AXToggle",
        "Calendar",
        False,
        2,
        True,
    )
    assert node.label() == "Calendar"


def test_prune_keeps_the_node_but_not_its_children() -> None:
    table = element("AXTable", element("AXRow"), AXDescription="Message List")
    rig = app_rig(window("Main", table))

    snapshot = rig.app.snapshot(table.parent or table, prune=lambda node: node.role == "AXTable")

    assert [node.role for node in snapshot.walk()] == ["AXWindow", "AXTable"]
