from __future__ import annotations

from collections.abc import Callable

from office_mcp.shared.errors import DuplicateSendError
from office_mcp.shared.text import normalize

DUPLICATE_WINDOW = 120.0


class TeamsSendGuard:
    def __init__(self, clock: Callable[[], float], window: float = DUPLICATE_WINDOW) -> None:
        self._clock = clock
        self._window = window
        self._sent: dict[tuple[str, str], float] = {}

    def check(self, thread_id: str, text: str) -> None:
        self._prune()
        sent_at = self._sent.get(key(thread_id, text))
        if sent_at is None:
            return
        age = self._clock() - sent_at
        raise DuplicateSendError(
            f"the same text was sent to this chat {age:.0f} s ago; nothing was sent. "
            "Pass allow_duplicate=true only when the owner wants it sent again."
        )

    def record(self, thread_id: str, text: str) -> None:
        self._prune()
        self._sent[key(thread_id, text)] = self._clock()

    def _prune(self) -> None:
        cutoff = self._clock() - self._window
        for stale in [entry for entry, sent_at in self._sent.items() if sent_at < cutoff]:
            del self._sent[stale]


def key(thread_id: str, text: str) -> tuple[str, str]:
    return thread_id, normalize(text)
