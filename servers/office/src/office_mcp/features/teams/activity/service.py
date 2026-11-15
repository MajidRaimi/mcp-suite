from __future__ import annotations

from dataclasses import dataclass

from office_mcp.features.teams.activity.models import ActivityFilter, ActivityItem
from office_mcp.features.teams.activity.parse import feed_ready, parse_activity
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.features.teams.reading.view import ReadScope
from office_mcp.shared.errors import ActionNotVerifiedError, ElementNotFoundError, UiTimeoutError
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

TOOLBAR_ID = "upfront-filter-toolbar-id"
FILTER_BUTTON = "Filter"
BUTTON_ROLE = "AXButton"
ON_VALUE = "1"
TOGGLE_WAIT = 3.0


@dataclass(frozen=True)
class FilterToggle:
    label: str
    wanted_for: ActivityFilter | None


TOGGLES = (
    FilterToggle("Unread", ActivityFilter.UNREAD),
    FilterToggle("@Mentions", ActivityFilter.MENTIONS),
    FilterToggle("Tag mentions", None),
)


def filter_toolbar(root: Node) -> Node | None:
    return root.find(lambda node: node.dom_id == TOOLBAR_ID)


def toggle_node(root: Node, label: str) -> Node | None:
    toolbar = filter_toolbar(root)
    if toolbar is None:
        return None
    wanted = normalize(label)
    return toolbar.find(
        lambda node: node is not toolbar and normalize(node.label()).startswith(wanted)
    )


def is_on(node: Node) -> bool:
    return node.value == ON_VALUE


def filter_button(root: Node) -> Node | None:
    return root.find(lambda node: node.role == BUTTON_ROLE and node.label() == FILTER_BUTTON)


class ActivityReader:
    def __init__(self, gateway: TeamsGateway, waiter: Waiter) -> None:
        self._gateway = gateway
        self._waiter = waiter

    def read(self, scope: ReadScope, wanted: ActivityFilter, limit: int) -> list[ActivityItem]:
        scope.show_app(TeamsApp.ACTIVITY)
        self._waiter.require(
            lambda: feed_ready(self._gateway.snapshot()),
            lambda: UiTimeoutError("the Teams activity feed did not load in time"),
        )
        for toggle in TOGGLES:
            self._apply(scope, toggle.label, toggle.wanted_for is wanted)
        items = self._settled()
        return items[:limit] if limit > 0 else items

    def _apply(self, scope: ReadScope, label: str, on: bool) -> None:
        node = self._toggle(label, required=on)
        if node is None or is_on(node) == on:
            return
        original = is_on(node)
        self._set(label, on)
        scope.cleanup.push(f"reset the activity filter {label}", lambda: self._set(label, original))

    def _toggle(self, label: str, required: bool) -> Node | None:
        root = self._gateway.snapshot()
        node = toggle_node(root, label)
        if node is not None or not required:
            return node
        button = filter_button(root)
        if button is not None and filter_toolbar(root) is None:
            self._gateway.press(button)
        return self._waiter.require(
            lambda: toggle_node(self._gateway.snapshot(), label),
            lambda: ElementNotFoundError(f"the Teams activity filter {label!r} was not found"),
            TOGGLE_WAIT,
        )

    def _set(self, label: str, on: bool) -> None:
        node = toggle_node(self._gateway.snapshot(), label)
        if node is None:
            raise ElementNotFoundError(f"the Teams activity filter {label!r} was not found")
        if is_on(node) == on:
            return
        self._gateway.press(node)
        self._waiter.require(
            lambda: self._shows(label, on),
            lambda: ActionNotVerifiedError(
                f"the Teams activity filter {label!r} did not switch {'on' if on else 'off'}"
            ),
            TOGGLE_WAIT,
        )

    def _shows(self, label: str, on: bool) -> bool:
        node = toggle_node(self._gateway.snapshot(), label)
        return node is not None and is_on(node) == on

    def _settled(self) -> list[ActivityItem]:
        latest: list[list[ActivityItem]] = [[]]

        def sample() -> int:
            latest[0] = parse_activity(self._gateway.snapshot())
            return len(latest[0])

        self._waiter.settle(sample)
        return latest[0]
