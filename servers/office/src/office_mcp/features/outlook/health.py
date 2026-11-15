from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Any

GO_ONLINE = "Go Online"
MIDNIGHT_GRACE = time(0, 10)
TODAY_BUTTON = re.compile(
    r"(?:^|,\s*)Today,\s*(?:[A-Za-z]+,?\s+)?"
    r"(?:(?P<day>\d{1,2})\.?\s+(?P<month>[A-Za-z]+)\.?"
    r"|(?P<month_first>[A-Za-z]+)\.?\s+(?P<day_last>\d{1,2}))"
    r"(?:,?\s+(?P<year>\d{4}))?(?:\s*,.*)?\s*$"
)
MONTHS = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
NOT_WORKING = "Outlook is not working right now."
OFFLINE_PROBLEM = (
    "Outlook is offline (it shows 'Outlook is offline' with a Go Online button), so no new mail "
    "is arriving and everything it shows is out of date. Sign in again or press Go Online in "
    "Outlook, then try again."
)


@dataclass(frozen=True)
class OutlookHealth:
    online: bool
    shown_today: date | None
    system_now: datetime

    @property
    def system_today(self) -> date:
        return self.system_now.date()

    @property
    def stale(self) -> bool:
        if self.shown_today is None or self.shown_today == self.system_today:
            return False
        just_after_midnight = self.system_now.time() < MIDNIGHT_GRACE
        return not (just_after_midnight and self.shown_today == self.system_today - timedelta(1))

    @property
    def problems(self) -> tuple[str, ...]:
        found: list[str] = []
        if not self.online:
            found.append(OFFLINE_PROBLEM)
        if self.stale and self.shown_today is not None:
            found.append(stale_problem(self.shown_today, self.system_today))
        return tuple(found)

    @property
    def working(self) -> bool:
        return not self.problems

    def message(self) -> str:
        return " ".join((NOT_WORKING, *self.problems))

    def to_record(self) -> dict[str, Any]:
        return {
            "outlook_working": self.working,
            "outlook_online": self.online,
            "outlook_today": self.shown_today.isoformat() if self.shown_today else None,
            "outlook_problems": list(self.problems),
        }


def unknown_health_record(problem: str) -> dict[str, Any]:
    return {
        "outlook_working": False,
        "outlook_online": None,
        "outlook_today": None,
        "outlook_problems": [problem],
    }


def stale_problem(shown: date, today: date) -> str:
    return (
        f"Outlook's view is stuck on {spoken(shown)} while today is {spoken(today)}, so its "
        "Today and Yesterday groups, times and new mail are wrong. Quit and reopen Outlook, then "
        "try again."
    )


def spoken(day: date) -> str:
    return f"{day:%A} {day.day} {day:%B} {day.year}"


def read_health(button_labels: Iterable[str], system_now: datetime) -> OutlookHealth:
    labels = list(button_labels)
    return OutlookHealth(
        online=GO_ONLINE not in labels,
        shown_today=shown_today(labels, system_now.date()),
        system_now=system_now,
    )


def shown_today(labels: Iterable[str], system_today: date) -> date | None:
    for label in labels:
        found = parse_today_label(label, system_today)
        if found is not None:
            return found
    return None


def parse_today_label(label: str, around: date) -> date | None:
    match = TODAY_BUTTON.search(label.strip())
    if not match:
        return None
    month_name = match["month"] or match["month_first"]
    day_text = match["day"] or match["day_last"]
    month = month_number(month_name)
    if month is None:
        return None
    if match["year"]:
        return safe_date(int(match["year"]), month, int(day_text))
    return nearest_date(month, int(day_text), around)


def month_number(name: str) -> int | None:
    prefix = name.casefold()[:3]
    return MONTHS.index(prefix) + 1 if prefix in MONTHS else None


def nearest_date(month: int, day: int, around: date) -> date | None:
    candidates = [
        found
        for year in (around.year - 1, around.year, around.year + 1)
        if (found := safe_date(year, month, day)) is not None
    ]
    return min(candidates, key=lambda found: abs((found - around).days), default=None)


def safe_date(year: int, month: int, day: int) -> date | None:
    try:
        return date(year, month, day)
    except ValueError:
        return None
