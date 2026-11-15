from __future__ import annotations

from collections.abc import Callable
from contextlib import AbstractContextManager

from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.scope import ScopeFactory, ScopeOptions, UiScope


class RecordingScope:
    def __init__(
        self,
        options: ScopeOptions,
        log: list[str],
        outcome: FocusOutcome = FocusOutcome.NOT_NEEDED,
    ) -> None:
        self.options = options
        self.cleanup = CleanupStack()
        self.log = log
        self._outcome = outcome

    def __enter__(self) -> RecordingScope:
        self.log.append(f"enter foreground={self.options.foreground}")
        self.cleanup.__enter__()
        return self

    def __exit__(self, *exc: object) -> None:
        error = exc[1] if len(exc) > 1 and isinstance(exc[1], BaseException) else None
        try:
            self.cleanup.__exit__(type(error) if error else None, error, None)
        finally:
            self.log.append("exit")

    def ensure_front(self) -> None:
        self.log.append("ensure_front")

    @property
    def focus(self) -> FocusOutcome:
        return self._outcome


def recording_factory(
    outcome: FocusOutcome = FocusOutcome.NOT_NEEDED,
) -> tuple[ScopeFactory, list[str]]:
    log: list[str] = []

    def factory(options: ScopeOptions) -> AbstractContextManager[UiScope]:
        return RecordingScope(options, log, outcome)

    return factory, log


def checked_factory(
    check: Callable[[], None], outcome: FocusOutcome = FocusOutcome.NOT_NEEDED
) -> tuple[ScopeFactory, list[str]]:
    log: list[str] = []

    def factory(options: ScopeOptions) -> AbstractContextManager[UiScope]:
        check()
        return RecordingScope(options, log, outcome)

    return factory, log
