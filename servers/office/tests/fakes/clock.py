from __future__ import annotations

from office_mcp.shared.waiting import Waiter


class FakeClock:
    def __init__(self, start: float = 0.0) -> None:
        self.now = start
        self.sleeps: list[float] = []

    def __call__(self) -> float:
        return self.now

    def sleep(self, seconds: float) -> None:
        self.sleeps.append(seconds)
        self.now += seconds

    def advance(self, seconds: float) -> None:
        self.now += seconds


def fake_waiter(clock: FakeClock | None = None, timeout: float = 12.0, poll: float = 0.4) -> Waiter:
    current = clock or FakeClock()
    return Waiter(clock=current, sleep=current.sleep, timeout=timeout, poll=poll)
