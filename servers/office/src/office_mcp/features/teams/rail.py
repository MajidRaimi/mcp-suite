from __future__ import annotations

from enum import StrEnum

from office_mcp.shared.node import Node

RAIL_ROLES = frozenset({"AXCheckBox", "AXToggleButton", "AXButton"})
SELECTED_VALUE = "1"


class TeamsApp(StrEnum):
    ACTIVITY = "14d6962d-6eeb-4f48-8890-de55454bb136"
    CHAT = "86fcd49b-61a2-4701-b771-54728cd291fb"
    TEAMS = "2a84919f-59d8-4441-a975-2a8c2643b741"
    CALENDAR = "ef56c0de-36fc-4ef8-b417-3d82ba9d073c"


def rail_button(root: Node, app: TeamsApp) -> Node | None:
    return root.find(lambda node: node.role in RAIL_ROLES and node.dom_id == app.value)


def selected_app(root: Node) -> TeamsApp | None:
    for app in TeamsApp:
        button = rail_button(root, app)
        if button is not None and button.value == SELECTED_VALUE:
            return app
    return None
