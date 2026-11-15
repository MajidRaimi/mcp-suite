from __future__ import annotations

import fcntl
import os
import threading
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.shared.errors import BusyError, OfficeError, error_text

READ_BUDGET = 120.0
WRITE_BUDGET = 30.0
SEND_BUDGET = 90.0
BATCH_BUDGET = 60.0
LOCK_POLL_SECONDS = 0.2
LOCK_DIRECTORY_MODE = 0o700
LOCK_FILE_MODE = 0o600
BUSY_MESSAGE = "another office action is still running; try again in a moment"


@dataclass(frozen=True)
class Deadline:
    clock: Callable[[], float]
    expires_at: float

    @classmethod
    def after(cls, clock: Callable[[], float], seconds: float) -> Deadline:
        return cls(clock=clock, expires_at=clock() + seconds)

    def remaining(self) -> float:
        return max(0.0, self.expires_at - self.clock())

    def expired(self) -> bool:
        return self.clock() >= self.expires_at


class UiLock(Protocol):
    def acquire(self, timeout: float) -> bool: ...

    def release(self) -> None: ...


class ProcessUiLock:
    def __init__(
        self, path: Path, clock: Callable[[], float], sleep: Callable[[float], None]
    ) -> None:
        self._path = path
        self._clock = clock
        self._sleep = sleep
        self._thread_lock = threading.Lock()
        self._descriptor: int | None = None

    def acquire(self, timeout: float) -> bool:
        deadline = self._clock() + timeout
        if not self._thread_lock.acquire(timeout=max(0.0, timeout)):
            return False
        try:
            descriptor = self._open()
        except OSError:
            self._thread_lock.release()
            raise
        while not self._try_flock(descriptor):
            remaining = deadline - self._clock()
            if remaining <= 0:
                os.close(descriptor)
                self._thread_lock.release()
                return False
            self._sleep(min(LOCK_POLL_SECONDS, remaining))
        self._descriptor = descriptor
        return True

    def release(self) -> None:
        descriptor, self._descriptor = self._descriptor, None
        if descriptor is None:
            return
        try:
            fcntl.flock(descriptor, fcntl.LOCK_UN)
        finally:
            os.close(descriptor)
            self._thread_lock.release()

    def _open(self) -> int:
        self._path.parent.mkdir(mode=LOCK_DIRECTORY_MODE, parents=True, exist_ok=True)
        return os.open(self._path, os.O_RDWR | os.O_CREAT, LOCK_FILE_MODE)

    @staticmethod
    def _try_flock(descriptor: int) -> bool:
        try:
            fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            return False
        return True


class Runner:
    def __init__(
        self, lock: UiLock, clock: Callable[[], float], lock_timeout: float, stale_after: float
    ) -> None:
        self._lock = lock
        self._clock = clock
        self._lock_timeout = lock_timeout
        self._stale_after = stale_after

    def read[T](self, work: Callable[[Deadline], T], budget: float = READ_BUDGET) -> T:
        return self._run(work, budget, refuse_stale=False)

    def write[T](self, work: Callable[[Deadline], T], budget: float = WRITE_BUDGET) -> T:
        return self._run(work, budget, refuse_stale=True)

    def unlocked[T](self, work: Callable[[], T]) -> T:
        try:
            return work()
        except OfficeError as error:
            raise ToolError(error_text(error)) from error

    def _run[T](self, work: Callable[[Deadline], T], budget: float, refuse_stale: bool) -> T:
        try:
            waited = self._acquire()
            try:
                if refuse_stale and waited > self._stale_after:
                    raise BusyError(stale_message(waited))
                return work(Deadline.after(self._clock, budget))
            finally:
                self._lock.release()
        except OfficeError as error:
            raise ToolError(error_text(error)) from error

    def _acquire(self) -> float:
        started = self._clock()
        if not self._lock.acquire(self._lock_timeout):
            raise BusyError(BUSY_MESSAGE)
        return self._clock() - started


def stale_message(waited: float) -> str:
    return (
        f"another office action held the apps for {waited:.0f} s, so nothing was done to "
        "avoid acting on a changed screen; try again"
    )


class Lazy[T]:
    def __init__(self, build: Callable[[], T]) -> None:
        self._build = build
        self._lock = threading.Lock()
        self._built: list[T] = []

    def get(self) -> T:
        with self._lock:
            if not self._built:
                self._built.append(self._build())
            return self._built[0]
