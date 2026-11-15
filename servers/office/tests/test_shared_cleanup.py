from __future__ import annotations

import pytest

from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.errors import (
    CleanupError,
    MenuStuckError,
    SendRefusedError,
    error_text,
)


def test_runs_in_lifo_order_and_reports() -> None:
    ran: list[str] = []

    with CleanupStack() as stack:
        stack.push("close menu", lambda: ran.append("menu"))
        stack.push("restore tab", lambda: ran.append("tab"))

    assert ran == ["tab", "menu"]
    assert stack.report == ("restore tab", "close menu")


def failing() -> None:
    raise MenuStuckError("menu stayed open")


def test_failures_become_notes_on_the_original_error() -> None:
    ran: list[str] = []

    with pytest.raises(SendRefusedError) as caught, CleanupStack() as stack:
        stack.push("restore view", lambda: ran.append("view"))
        stack.push("close menu", failing)
        raise SendRefusedError("nothing was sent")

    assert ran == ["view"]
    assert error_text(caught.value) == (
        "nothing was sent\ncleanup close menu failed: menu stayed open"
    )


def test_failure_after_success_raises_cleanup_error() -> None:
    ran: list[str] = []

    expected = "the action was done, but cleanup close menu failed"
    with pytest.raises(CleanupError, match=expected), CleanupStack() as stack:
        stack.push("restore view", lambda: ran.append("view"))
        stack.push("close menu", failing)

    assert ran == ["view"]
