from __future__ import annotations

from office_mcp.features.teams.ax import find_popup
from office_mcp.shared.node import Node


def menu(dom_id: str = "") -> Node:
    return Node(role="AXMenu", dom_id=dom_id, children=(Node(role="AXMenuItem", title="Copy"),))


def test_a_menu_inside_the_tracking_view_panel_is_not_a_popup() -> None:
    panel = Node(
        role="AXGroup",
        subrole="AXLandmarkRegion",
        dom_id="calv2-sf-tracking-view-focus",
        children=(menu(),),
    )
    root = Node(role="AXWindow", children=(Node(role="AXGroup", dom_id="app", children=(panel,)),))

    assert find_popup(root) is None


def test_a_menu_outside_persistent_panels_is_a_popup() -> None:
    open_menu = menu("context")
    panel = Node(role="AXGroup", dom_id="calv2-sf-tracking-view-focus", children=(menu(),))
    root = Node(role="AXWindow", children=(panel, open_menu))

    assert find_popup(root) is open_menu


def test_dialogs_are_popups() -> None:
    dialog = Node(role="AXGroup", subrole="AXApplicationDialog", title="Delete message?")

    assert find_popup(Node(role="AXWindow", children=(dialog,))) is dialog
