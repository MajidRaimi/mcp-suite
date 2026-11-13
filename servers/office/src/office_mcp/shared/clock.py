from __future__ import annotations

import re
from datetime import UTC, date, datetime, timedelta, tzinfo

WEEKDAYS = ("mon", "tue", "wed", "thu", "fri", "sat", "sun")
TIME_LABEL = re.compile(r"\d{1,2}:\d{2}(\s*[ap]\.?m\.?)?")
DATE_LABEL = re.compile(r"(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?")
EPOCH_MS_DIGITS = 13


def message_time(message_id: str, zone: tzinfo) -> datetime | None:
    if not message_id.isdigit() or len(message_id) != EPOCH_MS_DIGITS:
        return None
    return datetime.fromtimestamp(int(message_id) / 1000, tz=UTC).astimezone(zone)


def list_label_date(label: str, today: date) -> date | None:
    text = label.strip().casefold()
    if TIME_LABEL.fullmatch(text):
        return today
    if text == "yesterday":
        return today - timedelta(days=1)
    if text[:3] in WEEKDAYS and text.isalpha():
        days_back = (today.weekday() - WEEKDAYS.index(text[:3])) % 7 or 7
        return today - timedelta(days=days_back)
    match = DATE_LABEL.fullmatch(text)
    if not match:
        return None
    return calendar_date(int(match[1]), int(match[2]), match[3], today)


def calendar_date(month: int, day: int, year: str | None, today: date) -> date | None:
    try:
        if year:
            full_year = int(year) + 2000 if len(year) == 2 else int(year)
            return date(full_year, month, day)
        candidate = date(today.year, month, day)
        return candidate if candidate <= today else date(today.year - 1, month, day)
    except ValueError:
        return None
