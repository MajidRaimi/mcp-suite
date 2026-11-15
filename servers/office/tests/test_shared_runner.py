from __future__ import annotations

import stat
import time
from pathlib import Path

import pytest
from fakes.clock import FakeClock
from fakes.lock import MemoryUiLock
from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.shared.errors import ElementNotFoundError
from office_mcp.shared.runner import Deadline, Lazy, ProcessUiLock, Runner


def runner(lock: MemoryUiLock, clock: FakeClock) -> Runner:
    return Runner(lock, clock, lock_timeout=20.0, stale_after=10.0)


def test_read_runs_work_with_a_deadline_that_starts_after_the_lock() -> None:
    clock = FakeClock()
    lock = MemoryUiLock(clock, wait=5.0)

    deadline = runner(lock, clock).read(lambda d: d, budget=120.0)

    assert deadline.expires_at == pytest.approx(125.0)
    assert lock.timeouts == [20.0]
    assert lock.released == 1


def test_lock_timeout_raises_busy_tool_error() -> None:
    clock = FakeClock()
    lock = MemoryUiLock(clock)
    lock.available = False

    with pytest.raises(ToolError, match="another office action is still running"):
        runner(lock, clock).read(lambda d: "never")
    assert lock.released == 0


def test_write_refuses_after_a_stale_wait_and_does_nothing() -> None:
    clock = FakeClock()
    lock = MemoryUiLock(clock, wait=12.0)
    done: list[str] = []

    with pytest.raises(ToolError, match="nothing was done"):
        runner(lock, clock).write(lambda d: done.append("acted"))

    assert done == []
    assert lock.released == 1
    assert runner(MemoryUiLock(clock, wait=12.0), clock).read(lambda d: "read ok") == "read ok"


def test_office_errors_become_tool_errors_with_notes() -> None:
    clock = FakeClock()
    lock = MemoryUiLock(clock)

    def work(deadline: Deadline) -> None:
        error = ElementNotFoundError("row vanished")
        error.add_note("cleanup restore tab failed: tab missing")
        raise error

    with pytest.raises(ToolError) as caught:
        runner(lock, clock).write(work)

    assert str(caught.value) == "row vanished\ncleanup restore tab failed: tab missing"
    assert lock.released == 1


def test_deadline_remaining_and_expiry() -> None:
    clock = FakeClock()
    deadline = Deadline.after(clock, 30.0)

    clock.advance(20.0)
    assert deadline.remaining() == pytest.approx(10.0)
    assert not deadline.expired()
    clock.advance(15.0)
    assert deadline.remaining() == 0.0
    assert deadline.expired()


def test_process_lock_excludes_a_second_holder(tmp_path: Path) -> None:
    path = tmp_path / "cache" / "office-mcp" / "ui.lock"
    first = ProcessUiLock(path, time.monotonic, time.sleep)
    second = ProcessUiLock(path, time.monotonic, time.sleep)

    assert first.acquire(1.0)
    assert not second.acquire(0.3)
    first.release()
    assert second.acquire(1.0)
    second.release()

    assert stat.S_IMODE(path.stat().st_mode) == 0o600
    assert stat.S_IMODE(path.parent.stat().st_mode) == 0o700


def test_lazy_builds_once() -> None:
    built: list[int] = []

    def build() -> int:
        built.append(1)
        return 42

    lazy = Lazy(build)

    assert lazy.get() == 42
    assert lazy.get() == 42
    assert built == [1]
