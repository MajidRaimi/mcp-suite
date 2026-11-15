from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from enum import StrEnum
from typing import Any

from office_mcp.shared.drafts import Review
from office_mcp.shared.focus import FocusOutcome

WEEK_DAYS = 7
ONLINE_MEETING_MARKERS = ("microsoft teams meeting", "teams meeting")


class MyResponse(StrEnum):
    ACCEPTED = "accepted"
    TENTATIVE = "tentative"
    DECLINED = "declined"
    NOT_RESPONDED = "not_responded"
    ORGANIZER = "organizer"
    UNKNOWN = "unknown"


class RsvpResponse(StrEnum):
    ACCEPT = "accept"
    TENTATIVE = "tentative"
    DECLINE = "decline"

    @property
    def state(self) -> MyResponse:
        return RESPONSE_STATES[self]


RESPONSE_STATES: dict[RsvpResponse, MyResponse] = {
    RsvpResponse.ACCEPT: MyResponse.ACCEPTED,
    RsvpResponse.TENTATIVE: MyResponse.TENTATIVE,
    RsvpResponse.DECLINE: MyResponse.DECLINED,
}


@dataclass(frozen=True)
class DateRange:
    first: date
    last: date

    def contains(self, day: date) -> bool:
        return self.first <= day <= self.last

    def shifted(self, days: int) -> DateRange:
        step = timedelta(days=days)
        return DateRange(self.first + step, self.last + step)

    def describe(self) -> str:
        return f"{self.first.isoformat()} to {self.last.isoformat()}"


@dataclass(frozen=True)
class EventTitle:
    subject: str
    day: date
    start_minutes: int | None
    end_minutes: int | None
    location: str
    organizer: str
    show_as: str
    recurring: bool
    canceled: bool

    @property
    def all_day(self) -> bool:
        return self.start_minutes is None


@dataclass(frozen=True)
class CalendarEvent:
    id: str
    subject: str
    start: datetime
    end: datetime
    all_day: bool
    location: str
    organizer: str
    show_as: str
    recurring: bool
    canceled: bool

    @property
    def day(self) -> date:
        return self.start.date()

    @property
    def online_meeting(self) -> bool:
        folded = self.location.casefold()
        return any(marker in folded for marker in ONLINE_MEETING_MARKERS)

    def start_text(self) -> str:
        return self.start.date().isoformat() if self.all_day else self.start.isoformat()

    def end_text(self) -> str:
        return self.end.date().isoformat() if self.all_day else self.end.isoformat()

    def to_record(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "subject": self.subject,
            "start": self.start_text(),
            "end": self.end_text(),
            "all_day": self.all_day,
            "location": self.location,
            "online_meeting": self.online_meeting,
            "organizer": self.organizer,
            "show_as": self.show_as,
            "recurring": self.recurring,
            "canceled": self.canceled,
        }


@dataclass(frozen=True)
class Responses:
    accepted: int | None = None
    tentative: int | None = None
    declined: int | None = None
    not_responded: int | None = None

    def to_record(self) -> dict[str, int | None]:
        return {
            "accepted": self.accepted,
            "tentative": self.tentative,
            "declined": self.declined,
            "not_responded": self.not_responded,
        }


@dataclass(frozen=True)
class Peek:
    subject: str
    time_text: str
    recurrence: str | None
    location: str
    organizer: str
    invited_text: str
    my_response: MyResponse
    responses: Responses
    notes: str
    has_join_button: bool


@dataclass(frozen=True)
class EventDetails:
    event: CalendarEvent
    peek: Peek

    def to_record(self) -> dict[str, Any]:
        return {
            "id": self.event.id,
            "subject": self.event.subject or self.peek.subject,
            "start": self.event.start_text(),
            "end": self.event.end_text(),
            "recurrence": self.peek.recurrence,
            "location": self.peek.location or self.event.location,
            "organizer": self.peek.organizer or self.event.organizer,
            "invited_text": self.peek.invited_text,
            "my_response": self.peek.my_response.value,
            "responses": self.peek.responses.to_record(),
            "notes": self.peek.notes,
            "has_join_button": self.peek.has_join_button,
        }


@dataclass(frozen=True)
class EventQuery:
    first: date
    days: int
    include_canceled: bool
    limit: int

    @property
    def last(self) -> date:
        return self.first + timedelta(days=self.days - 1)

    def covers(self, event: CalendarEvent) -> bool:
        return self.first <= event.day <= self.last and (
            self.include_canceled or not event.canceled
        )


@dataclass(frozen=True)
class RsvpRequest:
    event_id: str
    response: RsvpResponse
    notify: bool
    comment: str
    dry_run: bool


@dataclass(frozen=True)
class RsvpOutcome:
    changed: bool
    verified: bool
    detail: str


@dataclass(frozen=True)
class RsvpResult:
    request: RsvpRequest
    outcome: RsvpOutcome
    review: Review | None
    focus: FocusOutcome

    @property
    def ok(self) -> bool:
        return self.outcome.verified or self.request.dry_run

    def to_record(self) -> dict[str, Any]:
        return {
            "ok": self.ok,
            "event": self.request.event_id,
            "response": self.request.response.value,
            "notified": self.request.notify and self.outcome.changed,
            "verified": self.outcome.verified,
            "detail": self.outcome.detail,
            "persona": persona_record(self.review),
            "focus_restored": self.focus.value,
        }


def persona_record(review: Review | None) -> dict[str, Any] | None:
    if review is None:
        return None
    return {
        "version": review.persona_version,
        "fixes": [fix.to_record() for fix in review.fixes],
        "warnings": list(review.warnings),
    }
