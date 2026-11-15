from __future__ import annotations

from collections.abc import Callable
from contextlib import suppress
from dataclasses import dataclass
from enum import StrEnum
from typing import Protocol

from office_mcp.shared.ax import AxApp, as_bool
from office_mcp.shared.errors import (
    ActivationError,
    BusyError,
    OfficeError,
    SessionUnavailableError,
)
from office_mcp.shared.waiting import Waiter

ACTIVATION_ATTEMPTS = 3
ACTIVATION_WAIT = 1.0
RESTORE_WAIT = 1.0
OWNER_BUSY = "the owner is using the Mac; try again shortly"
SCREEN_LOCKED = "the Mac screen is locked, so the apps cannot be used right now; try again later"
SECURE_INPUT = (
    "secure input is on (a password field has focus on the Mac), so keys cannot be sent; "
    "try again later"
)


class FrontHook:
    def __init__(self) -> None:
        self._bringers: list[Callable[[], None]] = []

    def bind(self, bring: Callable[[], None]) -> Callable[[], None]:
        self._bringers.append(bring)
        return lambda: self._bringers.remove(bring)

    def bring_front(self) -> None:
        if self._bringers:
            self._bringers[-1]()


class FocusOutcome(StrEnum):
    NOT_NEEDED = "not_needed"
    RESTORED = "restored"
    OWNER_MOVED = "owner_moved"
    KEPT = "kept"
    FAILED = "failed"


class Workspace(Protocol):
    def frontmost_pid(self) -> int | None: ...

    def activate(self, pid: int) -> None: ...

    def is_running(self, pid: int) -> bool: ...

    def app_name(self, pid: int) -> str: ...

    def idle_seconds(self) -> float: ...

    def screen_locked(self) -> bool: ...

    def secure_input(self) -> bool: ...


@dataclass(frozen=True)
class FocusPolicy:
    restore: bool
    idle_seconds: float
    idle_wait: float = 10.0


def require_session(workspace: Workspace) -> None:
    if workspace.screen_locked():
        raise SessionUnavailableError(SCREEN_LOCKED)
    if workspace.secure_input():
        raise SessionUnavailableError(SECURE_INPUT)


class ForegroundSession:
    def __init__(
        self, workspace: Workspace, app: AxApp, waiter: Waiter, policy: FocusPolicy
    ) -> None:
        self._workspace = workspace
        self._app = app
        self._waiter = waiter
        self._policy = policy
        self._pid: int | None = None
        self._previous: int | None = None
        self._outcome = FocusOutcome.NOT_NEEDED

    def __enter__(self) -> ForegroundSession:
        require_session(self._workspace)
        self._wait_for_idle()
        self._pid = self._app.require_pid()
        self._previous = self._workspace.frontmost_pid()
        if self._previous != self._pid:
            self._activate(self._pid)
        return self

    def __exit__(self, *exc: object) -> None:
        self._outcome = self._give_back()

    def ensure_front(self) -> None:
        pid = self._pid if self._pid is not None else self._app.require_pid()
        if self._workspace.frontmost_pid() != pid:
            self._activate(pid)

    @property
    def outcome(self) -> FocusOutcome:
        return self._outcome

    def _wait_for_idle(self) -> None:
        if self._policy.idle_seconds <= 0:
            return
        self._waiter.require(
            lambda: self._workspace.idle_seconds() >= self._policy.idle_seconds,
            lambda: BusyError(OWNER_BUSY),
            self._policy.idle_wait,
        )

    def _activate(self, pid: int) -> None:
        for _ in range(ACTIVATION_ATTEMPTS):
            self._request_front(pid)
            if self._waiter.until(lambda: self._workspace.frontmost_pid() == pid, ACTIVATION_WAIT):
                return
        raise ActivationError(f"could not bring {self._app.app_name} to the front")

    def _request_front(self, pid: int) -> None:
        with suppress(OfficeError):
            self._app.set_attribute(self._app.application(), "AXFrontmost", True)
        self._workspace.activate(pid)
        self._raise_main_window()

    def _raise_main_window(self) -> None:
        with suppress(OfficeError):
            windows = self._app.windows()
            main = next(
                (w for w in windows if as_bool(self._app.attribute(w, "AXMain")) is True),
                windows[0] if windows else None,
            )
            if main is not None:
                self._app.perform(main, "AXRaise", "main window")

    def _give_back(self) -> FocusOutcome:
        pid, previous = self._pid, self._previous
        if pid is None or previous == pid:
            return FocusOutcome.NOT_NEEDED
        if not self._policy.restore:
            return FocusOutcome.KEPT
        try:
            if self._workspace.frontmost_pid() != pid:
                return FocusOutcome.OWNER_MOVED
            if previous is None or not self._workspace.is_running(previous):
                return FocusOutcome.KEPT
            self._workspace.activate(previous)
            restored = self._waiter.until(
                lambda: self._workspace.frontmost_pid() == previous, RESTORE_WAIT
            )
        except Exception:
            return FocusOutcome.FAILED
        return FocusOutcome.RESTORED if restored else FocusOutcome.FAILED
