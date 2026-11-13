from __future__ import annotations

from typing import Protocol

from office_mcp.shared.node import Node


class TeamsGateway(Protocol):
    def snapshot(self) -> Node: ...

    def press(self, node: Node) -> None: ...

    def scroll_into_view(self, node: Node) -> None: ...
