from __future__ import annotations

from typing import Any

from office_mcp.shared.ax import AxApp
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.node import Node

TEAMS_BUNDLE = "com.microsoft.teams2"
TEAMS_APP_NAME = "Microsoft Teams"
SETTLE_TIMEOUT = 12.0
SETTLED_MIN_NODES = 200


class AxTeams:
    def __init__(self, app: AxApp | None = None) -> None:
        self._app = app or AxApp(TEAMS_BUNDLE, TEAMS_APP_NAME)
        self._settled = False

    def trusted(self) -> bool:
        return self._app.trusted()

    def running(self) -> bool:
        return self._app.running()

    def snapshot(self) -> Node:
        self._app.require_trust()
        if self._settled:
            return self._app.snapshot(self._window())
        node = self._settle()
        self._settled = True
        return node

    def press(self, node: Node) -> None:
        self._app.perform(self._handle(node), "AXPress", node.label())

    def scroll_into_view(self, node: Node) -> None:
        self._app.perform(self._handle(node), "AXScrollToVisible", node.label())

    def _window(self) -> Any:
        return self._app.window(lambda _title: True, web_content=True)

    def _settle(self) -> Node:
        deadline = self._app.clock() + SETTLE_TIMEOUT
        previous = -1
        while True:
            node = self._app.snapshot(self._window())
            size = sum(1 for _ in node.walk())
            if (size >= SETTLED_MIN_NODES and size == previous) or self._app.clock() > deadline:
                return node
            previous = size
            self._app.sleep(0.7)

    @staticmethod
    def _handle(node: Node) -> Any:
        if node.handle is None:
            raise OfficeError("cannot act on an element without a handle")
        return node.handle
