from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from office_mcp.shared.errors import CleanupError, OfficeError, error_text


@dataclass(frozen=True)
class CleanupStep:
    label: str
    undo: Callable[[], None]


class CleanupStack:
    def __init__(self) -> None:
        self._steps: list[CleanupStep] = []
        self._report: list[str] = []

    def push(self, label: str, undo: Callable[[], None]) -> None:
        self._steps.append(CleanupStep(label, undo))

    def __enter__(self) -> CleanupStack:
        return self

    def __exit__(
        self, kind: type[BaseException] | None, error: BaseException | None, tb: object
    ) -> None:
        failures = self._unwind()
        if not failures:
            return
        if error is not None:
            for failure in failures:
                error.add_note(failure)
            return
        raise CleanupError("the action was done, but " + "; ".join(failures))

    @property
    def report(self) -> tuple[str, ...]:
        return tuple(self._report)

    def _unwind(self) -> list[str]:
        failures: list[str] = []
        while self._steps:
            step = self._steps.pop()
            try:
                step.undo()
            except Exception as failure:
                failures.append(f"cleanup {step.label} failed: {describe(failure)}")
            else:
                self._report.append(step.label)
        return failures


def describe(failure: Exception) -> str:
    return error_text(failure) if isinstance(failure, OfficeError) else repr(failure)
