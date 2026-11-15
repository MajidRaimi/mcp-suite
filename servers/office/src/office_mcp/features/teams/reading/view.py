from __future__ import annotations

from collections.abc import Callable
from contextlib import AbstractContextManager
from typing import Protocol

from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.node import Node
from office_mcp.shared.scope import ScopeOptions


class ReadScope(Protocol):
    cleanup: CleanupStack

    def show_app(self, app: TeamsApp) -> Node: ...


ReadScopeFactory = Callable[[ScopeOptions], AbstractContextManager[ReadScope]]
