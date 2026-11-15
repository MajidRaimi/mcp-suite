from __future__ import annotations

import pytest
from fakes.clock import FakeClock, fake_waiter

from office_mcp.shared.errors import UiTimeoutError


def test_until_returns_the_first_ready_value() -> None:
    clock = FakeClock()
    values = iter([None, False, "ready"])

    result = fake_waiter(clock).until(lambda: next(values))

    assert result == "ready"
    assert clock.sleeps == [0.4, 0.4]


def test_until_gives_up_at_the_timeout() -> None:
    clock = FakeClock()

    assert fake_waiter(clock).until(lambda: None, timeout=1.0) is None
    assert clock.now == pytest.approx(1.0)


def test_require_raises_the_failure() -> None:
    waiter = fake_waiter(timeout=0.5)

    with pytest.raises(UiTimeoutError, match="never"):
        waiter.require(lambda: None, lambda: UiTimeoutError("never"))
    assert waiter.require(lambda: 3, lambda: UiTimeoutError("never")) == 3


def test_settle_waits_for_stable_counts_above_minimum() -> None:
    clock = FakeClock()
    samples = iter([10, 60, 120, 120, 500])

    assert fake_waiter(clock).settle(lambda: next(samples), minimum=50) == 120


def test_settle_returns_last_sample_on_timeout() -> None:
    clock = FakeClock()
    counter = iter(range(100))

    assert fake_waiter(clock).settle(lambda: next(counter), timeout=1.0) == 3


def test_with_timeout_copies() -> None:
    waiter = fake_waiter(timeout=12.0)

    assert waiter.with_timeout(3.0).timeout == 3.0
    assert waiter.timeout == 12.0
