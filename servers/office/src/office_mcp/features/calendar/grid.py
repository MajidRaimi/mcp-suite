from __future__ import annotations

from datetime import tzinfo

from office_mcp.features.calendar.models import CalendarEvent, DateRange
from office_mcp.features.calendar.parse import (
    build_event,
    parse_event_title,
    parse_range_label,
    parse_view_label,
)
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize

LANDMARK_MAIN = "AXLandmarkMain"
CALENDAR_VIEW_PREFIX = "calendar view"
VIEW_PICKER_PREFIX = "change the view of your calendar"
PREVIOUS_WEEK_PREFIX = "go to previous"
NEXT_WEEK_PREFIX = "go to next"
TODAY_PREFIX = "go to today"
WEEK_VIEW = "week"
BUTTON_ROLES = frozenset({"AXButton", "AXPopUpButton", "AXMenuButton"})
OPTION_ROLES = frozenset(
    {
        "AXMenuItem",
        "AXMenuItemRadio",
        "AXRadioButton",
        "AXButton",
        "AXCheckBox",
        "AXCell",
        "AXRow",
        "AXStaticText",
    }
)
MENU_ROLES = frozenset({"AXMenu", "AXList", "AXListBox"})


def folded_label(node: Node) -> str:
    return normalize(node.label())


def calendar_view(root: Node) -> Node | None:
    return root.find(
        lambda node: (
            LANDMARK_MAIN in (node.role, node.subrole)
            and folded_label(node).startswith(CALENDAR_VIEW_PREFIX)
        )
    )


def toolbar_button(root: Node, prefix: str) -> Node | None:
    return root.find(
        lambda node: (
            node.role in BUTTON_ROLES
            and folded_label(node).startswith(prefix)
            and parse_event_title(node.label()) is None
        )
    )


def view_picker(root: Node) -> Node | None:
    return toolbar_button(root, VIEW_PICKER_PREFIX)


def shown_view(root: Node) -> str | None:
    picker = view_picker(root)
    return parse_view_label(picker.label()) if picker is not None else None


NAVIGATION_PREFIXES = (PREVIOUS_WEEK_PREFIX, NEXT_WEEK_PREFIX, TODAY_PREFIX)


def is_navigation(node: Node) -> bool:
    return node.role in BUTTON_ROLES and folded_label(node).startswith(NAVIGATION_PREFIXES)


def shown_range(root: Node) -> DateRange | None:
    for node in root.walk():
        if is_navigation(node):
            continue
        label = node.label() if node.role in BUTTON_ROLES else node.value or node.title
        if not label or parse_event_title(label) is not None:
            continue
        found = parse_range_label(label)
        if found is not None:
            return found
    return None


def view_option(root: Node, view: str) -> Node | None:
    wanted = normalize(view)
    for menu in (node for node in root.walk() if node.role in MENU_ROLES):
        found = menu.find(lambda node: is_option(node, wanted))
        if found is not None:
            return found
    return None


def is_option(node: Node, wanted: str) -> bool:
    if node.role not in OPTION_ROLES:
        return False
    label = folded_label(node)
    return label == wanted or label.startswith(f"{wanted} ")


def event_buttons(root: Node, zone: tzinfo) -> list[tuple[CalendarEvent, Node]]:
    found: list[tuple[CalendarEvent, Node]] = []
    for node in root.walk():
        if node.role != "AXButton":
            continue
        event = build_event(node.title or node.label(), zone)
        if event is not None:
            found.append((event, node))
    return found


def buttons_for(root: Node, zone: tzinfo, event_id: str) -> list[tuple[CalendarEvent, Node]]:
    return [pair for pair in event_buttons(root, zone) if pair[0].id == event_id]
