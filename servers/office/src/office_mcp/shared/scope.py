from __future__ import annotations

from collections.abc import Callable
from contextlib import AbstractContextManager
from dataclasses import dataclass
from typing import Protocol

from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.focus import FocusOutcome


@dataclass(frozen=True)
class ScopeOptions:
    foreground: bool
    restore_view: bool = True


class UiScope(Protocol):
    cleanup: CleanupStack

    def ensure_front(self) -> None: ...

    @property
    def focus(self) -> FocusOutcome: ...


ScopeFactory = Callable[[ScopeOptions], AbstractContextManager[UiScope]]
