from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace

from office_mcp.shared.errors import OfficeError


@dataclass(frozen=True)
class Waiter:
    clock: Callable[[], float]
    sleep: Callable[[float], None]
    timeout: float = 12.0
    poll: float = 0.4

    def until[T](self, ready: Callable[[], T | None], timeout: float | None = None) -> T | None:
        deadline = self.clock() + (self.timeout if timeout is None else timeout)
        while True:
            result = ready()
            if result is not None and result is not False:
                return result
            remaining = deadline - self.clock()
            if remaining <= 0:
                return None
            self.sleep(min(self.poll, remaining))

    def require[T](
        self,
        ready: Callable[[], T | None],
        failure: Callable[[], OfficeError],
        timeout: float | None = None,
    ) -> T:
        result = self.until(ready, timeout)
        if result is None:
            raise failure()
        return result

    def settle(
        self,
        sample: Callable[[], int],
        minimum: int = 0,
        stable_rounds: int = 2,
        timeout: float | None = None,
    ) -> int:
        deadline = self.clock() + (self.timeout if timeout is None else timeout)
        previous = sample()
        rounds = 1
        while not (rounds >= stable_rounds and previous >= minimum):
            remaining = deadline - self.clock()
            if remaining <= 0:
                return previous
            self.sleep(min(self.poll, remaining))
            current = sample()
            rounds = rounds + 1 if current == previous else 1
            previous = current
        return previous

    def with_timeout(self, timeout: float) -> Waiter:
        return replace(self, timeout=timeout)
