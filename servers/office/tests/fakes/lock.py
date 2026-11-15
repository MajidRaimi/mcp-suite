from __future__ import annotations

from fakes.clock import FakeClock


class MemoryUiLock:
    def __init__(self, clock: FakeClock | None = None, wait: float = 0.0) -> None:
        self.clock = clock
        self.wait = wait
        self.available = True
        self.held = False
        self.acquired = 0
        self.released = 0
        self.timeouts: list[float] = []

    def acquire(self, timeout: float) -> bool:
        self.timeouts.append(timeout)
        if self.clock is not None:
            self.clock.advance(min(self.wait, timeout))
        if not self.available or self.held or self.wait > timeout:
            return False
        self.held = True
        self.acquired += 1
        return True

    def release(self) -> None:
        self.held = False
        self.released += 1
