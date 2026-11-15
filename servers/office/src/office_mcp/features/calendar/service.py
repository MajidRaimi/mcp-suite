from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, tzinfo

from office_mcp.features.calendar.cards import EventCards, pick_event
from office_mcp.features.calendar.gateway import (
    CalendarGateway,
    CalendarScope,
    CalendarScopeFactory,
)
from office_mcp.features.calendar.grid import buttons_for, event_buttons
from office_mcp.features.calendar.locations import EventLocations, require_event_id
from office_mcp.features.calendar.models import (
    CalendarEvent,
    EventDetails,
    EventQuery,
    RsvpOutcome,
    RsvpRequest,
    RsvpResult,
)
from office_mcp.features.calendar.navigation import CalendarNavigator, CalendarPage
from office_mcp.features.calendar.peek import parse_peek
from office_mcp.features.calendar.rsvp import Responder, RsvpPolicy
from office_mcp.shared.drafts import TextGuard
from office_mcp.shared.errors import AmbiguousTargetError, TargetError
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.node import Node
from office_mcp.shared.recipients import RecipientGuard
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import ScopeOptions
from office_mcp.shared.waiting import Waiter

MAX_DAYS = 31
MAX_LIMIT = 300
READ_OPTIONS = ScopeOptions(foreground=False)
RESPOND_OPTIONS = ScopeOptions(foreground=True)
DRY_RUN_DETAIL = "dry run: the checks passed and nothing was sent"


def event_query(
    start: str | None, days: int, include_canceled: bool, limit: int, today: date
) -> EventQuery:
    if not 1 <= days <= MAX_DAYS:
        raise TargetError(f"days must be between 1 and {MAX_DAYS}")
    if not 1 <= limit <= MAX_LIMIT:
        raise TargetError(f"limit must be between 1 and {MAX_LIMIT}")
    return EventQuery(parse_start(start, today), days, include_canceled, limit)


def parse_start(start: str | None, today: date) -> date:
    if start is None or not start.strip():
        return today
    try:
        return date.fromisoformat(start.strip())
    except ValueError as error:
        raise TargetError("start must be an ISO date such as 2026-11-15") from error


def select_events(events: list[CalendarEvent], query: EventQuery) -> list[CalendarEvent]:
    kept = sorted((event for event in events if query.covers(event)), key=sort_key)
    return kept[: query.limit]


def sort_key(event: CalendarEvent) -> tuple[str, bool, str, str]:
    return (event.day.isoformat(), not event.all_day, event.start.isoformat(), event.subject)


@dataclass(frozen=True)
class CalendarDeps:
    gateway: CalendarGateway
    scope: CalendarScopeFactory
    waiter: Waiter
    today: Callable[[], date]
    zone: tzinfo
    recipients: RecipientGuard
    guard: TextGuard


class CalendarService:
    def __init__(self, deps: CalendarDeps, locations: EventLocations | None = None) -> None:
        self._deps = deps
        self._locations = locations or EventLocations()
        self._cards = EventCards(deps.gateway, deps.waiter)
        self._policy = RsvpPolicy(deps.recipients, deps.guard)

    def events(
        self,
        start: str | None,
        days: int,
        include_canceled: bool,
        limit: int,
        deadline: Deadline,
    ) -> list[CalendarEvent]:
        query = event_query(start, days, include_canceled, limit, self._deps.today())
        with self._deps.scope(READ_OPTIONS) as scope:
            navigator = self._navigator(scope, deadline)
            page = navigator.show_week(navigator.open(), query.first)
            collected = self._collect(page)
            while page.range.last < query.last:
                page = navigator.next_week(page)
                collected.extend(self._collect(page))
        return select_events(collected, query)

    def event(self, event: str, deadline: Deadline) -> EventDetails:
        event_id = require_event_id(event)
        day = self._locations.day_of(event_id)
        with self._deps.scope(READ_OPTIONS) as scope:
            page = self._week_of(scope, deadline, day)
            found, button = pick_event(page.root, self._deps.zone, event_id)
            peek = parse_peek(self._cards.open(scope.cleanup, button))
        return EventDetails(found, peek)

    def respond(self, request: RsvpRequest, deadline: Deadline) -> RsvpResult:
        review = self._policy.review(request)
        day = self._locations.day_of(request.event_id)
        if request.dry_run:
            dry = RsvpOutcome(changed=False, verified=False, detail=DRY_RUN_DETAIL)
            return RsvpResult(request, dry, review, FocusOutcome.NOT_NEEDED)
        with self._deps.scope(RESPOND_OPTIONS) as scope:
            page = self._week_of(scope, deadline, day)
            _found, button = pick_event(page.root, self._deps.zone, request.event_id)
            peek = self._cards.open(scope.cleanup, button)
            responder = Responder(
                self._deps.gateway,
                self._deps.waiter,
                self._cards,
                self._deps.guard,
                lambda: self._reopen(scope, request.event_id),
            )
            outcome = responder.respond(scope.cleanup, peek, request, review)
        return RsvpResult(request, outcome, review, scope.focus)

    def _navigator(self, scope: CalendarScope, deadline: Deadline) -> CalendarNavigator:
        deps = self._deps
        return CalendarNavigator(deps.gateway, deps.waiter, scope, deadline, deps.today)

    def _week_of(self, scope: CalendarScope, deadline: Deadline, day: date) -> CalendarPage:
        navigator = self._navigator(scope, deadline)
        return navigator.show_week(navigator.open(), day)

    def _collect(self, page: CalendarPage) -> list[CalendarEvent]:
        events = [event for event, _node in event_buttons(page.root, self._deps.zone)]
        self._locations.remember(events)
        return events

    def _reopen(self, scope: CalendarScope, event_id: str) -> Node | None:
        self._cards.close()
        matches = buttons_for(self._deps.gateway.snapshot(), self._deps.zone, event_id)
        if not matches:
            return None
        if len(matches) > 1:
            raise AmbiguousTargetError("several events now share this id", [])
        return self._cards.open(scope.cleanup, matches[0][1])
