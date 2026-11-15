from __future__ import annotations

import re

from office_mcp.features.teams.activity.models import ActivityItem, ActivityKind
from office_mcp.shared.node import Node

FEED_LABEL = "Activity"
FEED_HEADER_ID = "feed-header"
LIST_ROLE = "AXList"
TITLE_PREFIX = "activity-feed-item-title-"
PREVIEW_PREFIX = "activity-feed-item-message-preview-"
TIME_PREFIX = "activity-feed-item-timestamp-"
LOCATION_PREFIX = "activity-feed-item-location-"
MISSED_CALL = re.compile(r"^missed (?:audio |video )?call from (?P<person>.+)$", re.IGNORECASE)
KIND_MARKERS: tuple[tuple[ActivityKind, re.Pattern[str]], ...] = (
    (ActivityKind.MENTION, re.compile(r"\bmentioned\b", re.IGNORECASE)),
    (ActivityKind.REPLY, re.compile(r"\breplied\b", re.IGNORECASE)),
    (ActivityKind.REACTION, re.compile(r"\breacted\b", re.IGNORECASE)),
    (ActivityKind.INVITE, re.compile(r"\binvited you\b", re.IGNORECASE)),
    (
        ActivityKind.MEETING,
        re.compile(r"\b(?:forwarded you|updated|canceled|cancelled|scheduled)\b", re.IGNORECASE),
    ),
)


def feed_list(root: Node) -> Node | None:
    return root.find(lambda node: node.role == LIST_ROLE and node.label() == FEED_LABEL)


def feed_ready(root: Node) -> bool:
    if feed_list(root) is not None:
        return True
    return root.find(is_feed_marker) is not None


def is_feed_marker(node: Node) -> bool:
    return node.dom_id == FEED_HEADER_ID or node.dom_id.startswith(TITLE_PREFIX)


def parse_activity(root: Node) -> list[ActivityItem]:
    scope = feed_list(root) or root
    by_id = {node.dom_id: node for node in scope.walk() if node.dom_id}
    keys = [
        node.dom_id.removeprefix(TITLE_PREFIX)
        for node in scope.walk()
        if node.dom_id.startswith(TITLE_PREFIX)
    ]
    return [build_item(key, by_id) for key in dict.fromkeys(keys)]


def build_item(key: str, by_id: dict[str, Node]) -> ActivityItem:
    title = text_of(by_id, TITLE_PREFIX + key)
    kind, person = classify(title)
    return ActivityItem(
        kind=kind,
        person=person,
        title=title,
        preview=text_of(by_id, PREVIEW_PREFIX + key),
        time_label=text_of(by_id, TIME_PREFIX + key),
        location=text_of(by_id, LOCATION_PREFIX + key),
    )


def text_of(by_id: dict[str, Node], dom_id: str) -> str:
    node = by_id.get(dom_id)
    return node.label() if node is not None else ""


def classify(title: str) -> tuple[ActivityKind, str]:
    missed = MISSED_CALL.match(title)
    if missed:
        return ActivityKind.MISSED_CALL, missed["person"].strip()
    for kind, marker in KIND_MARKERS:
        found = marker.search(title)
        if found:
            return kind, title[: found.start()].strip()
    return ActivityKind.OTHER, ""
