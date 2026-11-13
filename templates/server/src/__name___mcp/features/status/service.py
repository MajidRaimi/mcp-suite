from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, datetime, tzinfo

Clock = Callable[[], datetime]


def utc_now() -> datetime:
    return datetime.now(UTC)


@dataclass(frozen=True)
class StatusReport:
    owner: str
    timezone: str
    local_time: str


class StatusService:
    def __init__(self, owner_label: str, timezone: tzinfo, clock: Clock = utc_now) -> None:
        self._owner_label = owner_label
        self._timezone = timezone
        self._clock = clock

    def report(self) -> StatusReport:
        local = self._clock().astimezone(self._timezone)
        return StatusReport(
            owner=self._owner_label,
            timezone=str(self._timezone),
            local_time=local.isoformat(timespec="seconds"),
        )
