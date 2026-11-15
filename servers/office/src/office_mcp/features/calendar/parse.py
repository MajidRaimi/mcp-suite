from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, tzinfo

from office_mcp.features.calendar.models import CalendarEvent, DateRange, EventTitle

SEPARATOR = ", "
ORGANIZER_PREFIX = "By "
RECURRING_MARK = "recurring event"
CANCELED_PREFIXES = ("Canceled: ", "Cancelled: ")
ID_LENGTH = 12
NOON = 12
SPACE_LIKE = str.maketrans({"\u00a0": " ", "\u202f": " ", "\u2009": " "})
DASH_LIKE = str.maketrans({"\u2012": "-", "\u2013": "-", "\u2014": "-", "\u2212": "-"})
ALL_DAY_MARKS = frozenset({"all day", "all day event", "all-day", "all-day event"})
SHOW_AS: dict[str, str] = {
    "free": "free",
    "tentative": "tentative",
    "busy": "busy",
    "away": "out_of_office",
    "out of office": "out_of_office",
    "working elsewhere": "working_elsewhere",
}
MONTH_NAMES = (
    "january",
    "february",
    "march",
    "april",
    "may",
    "june",
    "july",
    "august",
    "september",
    "october",
    "november",
    "december",
)
MONTHS: dict[str, int] = {
    **{name: number for number, name in enumerate(MONTH_NAMES, start=1)},
    **{name[:3]: number for number, name in enumerate(MONTH_NAMES, start=1)},
    "sept": 9,
}
WEEKDAYS = frozenset(
    name
    for full in ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")
    for name in (full, full[:3])
)
CLOCK = r"(?P<{hour}>\d{{1,2}}):(?P<{minute}>\d{{2}})(?:\s*(?P<{half}>[AaPp][Mm]))?"
TIME_RANGE = re.compile(
    CLOCK.format(hour="sh", minute="sm", half="sp")
    + r"\s+to\s+"
    + CLOCK.format(hour="eh", minute="em", half="ep")
)
MONTH_DAY = re.compile(r"(?P<month>[A-Za-z]{3,9})\.?\s+(?P<day>\d{1,2})")
YEAR = re.compile(r"\d{4}")
MONTH = r"(?P<{name}>[A-Za-z]{{3,9}})\.?"
RANGE_PATTERNS = (
    re.compile(
        MONTH.format(name="m1")
        + r"\s+(?P<d1>\d{1,2}),\s*(?P<y1>\d{4})\s*-\s*"
        + MONTH.format(name="m2")
        + r"\s+(?P<d2>\d{1,2}),\s*(?P<y2>\d{4})"
    ),
    re.compile(
        MONTH.format(name="m1")
        + r"\s+(?P<d1>\d{1,2})\s*-\s*"
        + MONTH.format(name="m2")
        + r"\s+(?P<d2>\d{1,2}),\s*(?P<y2>\d{4})"
    ),
    re.compile(
        MONTH.format(name="m1") + r"\s+(?P<d1>\d{1,2})\s*-\s*(?P<d2>\d{1,2}),\s*(?P<y2>\d{4})"
    ),
)
VIEW_LABEL = re.compile(r"(?P<view>[A-Za-z][A-Za-z ]*?)\s+selected\.?\s*$")


def clean(text: str) -> str:
    return " ".join(text.translate(SPACE_LIKE).split())


def parse_event_title(title: str) -> EventTitle | None:
    tokens = clean(title).split(SEPARATOR)
    end = len(tokens)
    recurring = end > 0 and tokens[end - 1].casefold() == RECURRING_MARK
    end -= int(recurring)
    show_as = SHOW_AS.get(tokens[end - 1].casefold(), "") if end > 0 else ""
    end -= int(bool(show_as))
    anchor = date_anchor(tokens, end)
    if anchor is None:
        return None
    organizer_at = organizer_index(tokens, anchor.index + 4, end)
    location_end = end if organizer_at is None else organizer_at
    subject, canceled = split_canceled(SEPARATOR.join(tokens[: anchor.index]))
    return EventTitle(
        subject=subject,
        day=anchor.day,
        start_minutes=anchor.start_minutes,
        end_minutes=anchor.end_minutes,
        location=SEPARATOR.join(tokens[anchor.index + 4 : location_end]),
        organizer=organizer_name(tokens, organizer_at, end),
        show_as=show_as,
        recurring=recurring,
        canceled=canceled,
    )


@dataclass(frozen=True)
class DateAnchor:
    index: int
    day: date
    start_minutes: int | None
    end_minutes: int | None


def date_anchor(tokens: list[str], end: int) -> DateAnchor | None:
    for index in range(end - 4, -1, -1):
        span = time_span(tokens[index])
        if span is None or tokens[index + 1].casefold() not in WEEKDAYS:
            continue
        day = token_date(tokens[index + 2], tokens[index + 3])
        if day is not None:
            return DateAnchor(index, day, span[0], span[1])
    return None


def organizer_index(tokens: list[str], start: int, end: int) -> int | None:
    found = [index for index in range(start, end) if tokens[index].startswith(ORGANIZER_PREFIX)]
    return found[-1] if found else None


def organizer_name(tokens: list[str], start: int | None, end: int) -> str:
    if start is None:
        return ""
    return SEPARATOR.join(tokens[start:end])[len(ORGANIZER_PREFIX) :].strip()


def split_canceled(subject: str) -> tuple[str, bool]:
    for prefix in CANCELED_PREFIXES:
        if subject.startswith(prefix):
            return subject[len(prefix) :], True
    return subject, False


def time_span(token: str) -> tuple[int | None, int | None] | None:
    if token.casefold() in ALL_DAY_MARKS:
        return None, None
    match = TIME_RANGE.fullmatch(token)
    if match is None:
        return None
    start = clock_minutes(match["sh"], match["sm"], match["sp"] or match["ep"])
    end = clock_minutes(match["eh"], match["em"], match["ep"] or match["sp"])
    if start is None or end is None:
        return None
    return start, end


def clock_minutes(hour_text: str, minute_text: str, half: str | None) -> int | None:
    hour, minute = int(hour_text), int(minute_text)
    if minute > 59:
        return None
    if half is None:
        return hour * 60 + minute if hour < 24 else None
    if not 1 <= hour <= NOON:
        return None
    hour = hour % NOON + (NOON if half.casefold() == "pm" else 0)
    return hour * 60 + minute


def token_date(month_day: str, year: str) -> date | None:
    match = MONTH_DAY.fullmatch(month_day)
    if match is None or YEAR.fullmatch(year) is None:
        return None
    return safe_date(int(year), match["month"], int(match["day"]))


def safe_date(year: int, month_name: str, day: int) -> date | None:
    month = MONTHS.get(month_name.casefold())
    if month is None:
        return None
    try:
        return date(year, month, day)
    except ValueError:
        return None


def event_id(subject: str, start: str, organizer: str) -> str:
    digest = hashlib.sha1("\n".join((subject, start, organizer)).encode(), usedforsecurity=False)
    return digest.hexdigest()[:ID_LENGTH]


def build_event(title: str, zone: tzinfo) -> CalendarEvent | None:
    parsed = parse_event_title(title)
    if parsed is None:
        return None
    start = at_minutes(parsed.day, parsed.start_minutes or 0, zone)
    end = event_end(parsed, start, zone)
    displayed = displayed_subject(parsed)
    start_key = parsed.day.isoformat() if parsed.all_day else start.isoformat()
    return CalendarEvent(
        id=event_id(displayed, start_key, parsed.organizer),
        subject=parsed.subject,
        start=start,
        end=end,
        all_day=parsed.all_day,
        location=parsed.location,
        organizer=parsed.organizer,
        show_as=parsed.show_as,
        recurring=parsed.recurring,
        canceled=parsed.canceled,
    )


def displayed_subject(parsed: EventTitle) -> str:
    return f"{CANCELED_PREFIXES[0]}{parsed.subject}" if parsed.canceled else parsed.subject


def event_end(parsed: EventTitle, start: datetime, zone: tzinfo) -> datetime:
    if parsed.all_day or parsed.end_minutes is None:
        return at_minutes(parsed.day + timedelta(days=1), 0, zone)
    end = at_minutes(parsed.day, parsed.end_minutes, zone)
    return end if end > start else end + timedelta(days=1)


def at_minutes(day: date, minutes: int, zone: tzinfo) -> datetime:
    clock = time(minutes // 60 % 24, minutes % 60)
    return datetime.combine(day, clock, tzinfo=zone)


def parse_range_label(label: str) -> DateRange | None:
    text = clean(label).translate(DASH_LIKE)
    for pattern in RANGE_PATTERNS:
        for match in pattern.finditer(text):
            found = range_from(match.groupdict())
            if found is not None:
                return found
    return None


def range_from(groups: dict[str, str | None]) -> DateRange | None:
    year_last = int(groups["y2"] or 0)
    month_first = groups["m1"] or ""
    month_last = groups.get("m2") or month_first
    year_first = int(groups.get("y1") or year_last)
    last = safe_date(year_last, month_last, int(groups["d2"] or 0))
    first = safe_date(year_first, month_first, int(groups["d1"] or 0))
    if first is not None and last is not None and first > last and groups.get("y1") is None:
        first = safe_date(year_last - 1, month_first, first.day)
    if first is None or last is None or first > last:
        return None
    return DateRange(first, last)


def parse_view_label(label: str) -> str | None:
    match = VIEW_LABEL.search(clean(label))
    return match["view"].strip().casefold() if match is not None else None
