from __future__ import annotations

from office_mcp.shared.ax import AxApp
from office_mcp.shared.errors import CleanupError
from office_mcp.shared.events import InputEvents, Key, Shortcut
from office_mcp.shared.waiting import Waiter

SUGGESTIONS_WINDOW = "Search Suggestions"
APPEAR_SECONDS = 1.0
CLOSE_SECONDS = 2.0
CLOSE_ATTEMPTS = 2


def suggestions_open(app: AxApp) -> bool:
    return any(app.text(window, "AXDescription") == SUGGESTIONS_WINDOW for window in app.windows())


def await_suggestions(app: AxApp, waiter: Waiter) -> bool:
    return waiter.until(lambda: suggestions_open(app), APPEAR_SECONDS) is True


def close_suggestions(app: AxApp, events: InputEvents, waiter: Waiter) -> int:
    if not suggestions_open(app):
        return 0
    for _ in range(CLOSE_ATTEMPTS):
        events.press(app.require_pid(), Shortcut(Key.ESCAPE))
        if waiter.until(lambda: not suggestions_open(app), CLOSE_SECONDS):
            return 1
    raise CleanupError("the Outlook search suggestions list stayed open")
