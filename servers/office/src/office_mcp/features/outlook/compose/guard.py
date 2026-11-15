from __future__ import annotations

import hashlib
import threading
from collections.abc import Callable, Sequence
from dataclasses import dataclass

from office_mcp.shared.drafts import canonical_text
from office_mcp.shared.errors import DuplicateSendError
from office_mcp.shared.recipients import normalize_email
from office_mcp.shared.text import normalize

DUPLICATE_WINDOW = 120.0


@dataclass(frozen=True)
class SentRecord:
    fingerprint: str
    subject: str
    at: float


def send_fingerprint(action: str, targets: Sequence[str], subject: str, body: str) -> str:
    people = ",".join(sorted(normalize_email(target) for target in targets))
    text = "\n".join((action, people, normalize(subject), canonical_text(body)))
    return hashlib.sha256(text.encode()).hexdigest()


class SendGuard:
    def __init__(self, clock: Callable[[], float], window: float = DUPLICATE_WINDOW) -> None:
        self._clock = clock
        self._window = window
        self._lock = threading.Lock()
        self._records: list[SentRecord] = []

    def check(self, fingerprint: str) -> None:
        with self._lock:
            self._expire()
            if any(record.fingerprint == fingerprint for record in self._records):
                raise DuplicateSendError(
                    f"the same mail was sent less than {self._window:.0f} s ago; pass "
                    "allow_duplicate=true only if it really must go out twice"
                )

    def record(self, fingerprint: str, subject: str) -> SentRecord:
        entry = SentRecord(fingerprint=fingerprint, subject=subject, at=self._clock())
        with self._lock:
            self._expire()
            self._records.append(entry)
        return entry

    def last(self, within: float) -> SentRecord | None:
        with self._lock:
            if not self._records:
                return None
            newest = self._records[-1]
            return newest if self._clock() - newest.at <= within else None

    def forget(self, entry: SentRecord) -> None:
        with self._lock:
            self._records = [record for record in self._records if record is not entry]

    def _expire(self) -> None:
        now = self._clock()
        self._records = [record for record in self._records if now - record.at < self._window]
