from __future__ import annotations

from collections.abc import Callable, Sequence

from office_mcp.shared.ax import MAX_DEPTH, MAX_NODES, AxApp, as_bool, as_int, as_sequence, as_text
from office_mcp.shared.node import Node

SNAPSHOT_ATTRIBUTES = (
    "AXRole",
    "AXSubrole",
    "AXTitle",
    "AXValue",
    "AXDescription",
    "AXHelp",
    "AXEnabled",
    "AXDOMIdentifier",
    "AXIdentifier",
    "AXDOMClassList",
    "AXDisclosureLevel",
    "AXSelected",
    "AXChildren",
)


class Snapshotter:
    def __init__(
        self,
        app: AxApp,
        max_depth: int = MAX_DEPTH,
        prune: Callable[[Node], bool] | None = None,
        budget: int = MAX_NODES,
    ) -> None:
        self._app = app
        self._max_depth = max_depth
        self._prune = prune
        self._budget = budget

    def build(self, element: object, depth: int = 0) -> Node:
        self._budget -= 1
        values = self._app.attributes(element, SNAPSHOT_ATTRIBUTES)
        node = node_from_values(values, element, ())
        if depth >= self._max_depth or self._budget <= 0 or (self._prune and self._prune(node)):
            return node
        children = [
            child for child in as_sequence(values[-1]) if not self._app.same(child, element)
        ]
        kids: list[Node] = []
        for child in children:
            if self._budget <= 0:
                break
            kids.append(self.build(child, depth + 1))
        return node_from_values(values, element, tuple(kids))


def node_from_values(
    values: Sequence[object | None], element: object, children: tuple[Node, ...]
) -> Node:
    (
        role,
        subrole,
        title,
        value,
        description,
        help_text,
        enabled,
        dom_id,
        identifier,
        classes,
        level,
        selected,
        _,
    ) = values
    return Node(
        role=as_text(role),
        title=as_text(title),
        value=as_text(value),
        description=as_text(description),
        dom_id=as_text(dom_id) or as_text(identifier),
        classes=tuple(as_text(item) for item in as_sequence(classes)),
        children=children,
        handle=element,
        subrole=as_text(subrole),
        help=as_text(help_text),
        enabled=as_bool(enabled),
        level=as_int(level),
        selected=as_bool(selected),
    )
