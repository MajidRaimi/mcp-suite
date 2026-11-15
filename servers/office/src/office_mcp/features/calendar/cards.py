from __future__ import annotations

from datetime import tzinfo

from office_mcp.features.calendar.gateway import CalendarGateway
from office_mcp.features.calendar.grid import buttons_for
from office_mcp.features.calendar.models import CalendarEvent
from office_mcp.features.calendar.peek import find_peek
from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.errors import AmbiguousTargetError, TargetNotFoundError, UiTimeoutError
from office_mcp.shared.node import Node
from office_mcp.shared.waiting import Waiter

PEEK_WAIT = 6.0
CLOSE_CARD = "close the Teams event card"
GONE = "the event is no longer in the Teams calendar; list events again"
COLLISION = (
    "several events share this id (same subject, start and organizer, for example a "
    "duplicate invite); open it in Teams instead"
)


def pick_event(root: Node, zone: tzinfo, event_id: str) -> tuple[CalendarEvent, Node]:
    matches = buttons_for(root, zone, event_id)
    if not matches:
        raise TargetNotFoundError(GONE)
    if len(matches) > 1:
        raise AmbiguousTargetError(COLLISION, [event.subject for event, _node in matches])
    return matches[0]


class EventCards:
    def __init__(self, gateway: CalendarGateway, waiter: Waiter) -> None:
        self._gateway = gateway
        self._waiter = waiter

    def open(self, cleanup: CleanupStack, button: Node) -> Node:
        self._gateway.press(button)
        cleanup.push(CLOSE_CARD, self.close)
        return self._waiter.require(
            lambda: find_peek(self._gateway.snapshot()),
            lambda: UiTimeoutError("the Teams event card did not open"),
            PEEK_WAIT,
        )

    def current(self) -> Node | None:
        return find_peek(self._gateway.snapshot())

    def close(self) -> None:
        self._gateway.close_popups()
