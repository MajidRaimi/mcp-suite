from __future__ import annotations

import re
import threading
from collections import OrderedDict
from collections.abc import Iterable
from datetime import date

from office_mcp.features.calendar.models import CalendarEvent
from office_mcp.shared.errors import TargetError, TargetNotFoundError

EVENT_ID = re.compile(r"[0-9a-f]{12}")
ID_HINT = "use the id from calendar_events"
UNKNOWN_ID = (
    "this event id is not known (ids are kept until the server restarts); "
    "call calendar_events for that day first"
)
DEFAULT_CAPACITY = 2000


def require_event_id(text: str) -> str:
    candidate = text.strip()
    if EVENT_ID.fullmatch(candidate) is None:
        raise TargetError(ID_HINT)
    return candidate


class EventLocations:
    def __init__(self, capacity: int = DEFAULT_CAPACITY) -> None:
        self._capacity = capacity
        self._days: OrderedDict[str, date] = OrderedDict()
        self._lock = threading.Lock()

    def remember(self, events: Iterable[CalendarEvent]) -> None:
        with self._lock:
            for event in events:
                self._days[event.id] = event.day
                self._days.move_to_end(event.id)
            while len(self._days) > self._capacity:
                self._days.popitem(last=False)

    def day_of(self, event_id: str) -> date:
        with self._lock:
            day = self._days.get(event_id)
        if day is None:
            raise TargetNotFoundError(UNKNOWN_ID)
        return day
