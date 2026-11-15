from __future__ import annotations

import re
from enum import StrEnum

from office_mcp.shared.recipients import is_email
from office_mcp.shared.windows import WindowProbe

STANDARD_WINDOW = "AXStandardWindow"
MODULE_LANDMARKS = frozenset({"Mail", "Calendar"})
SEND_LANDMARKS = frozenset({"Send", "Send the message"})
MESSAGE_LIST = "Message List"
CANCEL_SEARCH = "Cancel Search"
CALENDAR_VIEW = "Calendar View"
SEARCHING_PREFIX = "Searching"
CALENDAR_TITLE = "Calendar"
TITLE_SEPARATOR = " • "
SEARCHING_TITLE = re.compile(r'^Searching\s+"(?P<account>[^"]+)"')


class OutlookWindow(StrEnum):
    MAIN = "main"
    COMPOSE = "compose"
    OTHER = "other"


class OutlookMode(StrEnum):
    MAIL = "mail"
    SEARCH = "search"
    CALENDAR = "calendar"
    OTHER = "other"


class OutlookModule(StrEnum):
    MAIL = "Mail"
    CALENDAR = "Calendar"


class FocusedTab(StrEnum):
    FOCUSED = "focused"
    OTHER = "other"


def has_module_switcher(probe: WindowProbe) -> bool:
    return probe.landmarks >= MODULE_LANDMARKS


def classify_window(probe: WindowProbe) -> OutlookWindow:
    if has_module_switcher(probe):
        return OutlookWindow.MAIN
    if probe.subrole == STANDARD_WINDOW and probe.landmarks & SEND_LANDMARKS:
        return OutlookWindow.COMPOSE
    return OutlookWindow.OTHER


def detect_mode(probe: WindowProbe) -> OutlookMode:
    if probe.title.startswith(SEARCHING_PREFIX) or CANCEL_SEARCH in probe.landmarks:
        return OutlookMode.SEARCH
    if probe.title == CALENDAR_TITLE or any(
        label.startswith(CALENDAR_VIEW) for label in probe.landmarks
    ):
        return OutlookMode.CALENDAR
    if MESSAGE_LIST in probe.landmarks or TITLE_SEPARATOR in probe.title:
        return OutlookMode.MAIL
    return OutlookMode.OTHER


def account_from_title(title: str) -> str | None:
    searching = SEARCHING_TITLE.match(title.strip())
    if searching:
        candidate = searching["account"].strip()
    else:
        _, separator, candidate = title.rpartition(TITLE_SEPARATOR)
        candidate = candidate.strip() if separator else ""
    return candidate if is_email(candidate) else None


def folder_from_title(title: str) -> str:
    folder, separator, _ = title.partition(TITLE_SEPARATOR)
    return folder.strip() if separator else ""
