from __future__ import annotations

import hashlib
import re
from dataclasses import replace
from datetime import date, timedelta

from office_mcp.features.outlook.models import Folder, Row, Summary
from office_mcp.shared.node import Node

FOLDER = re.compile(r"^(?P<name>.+?)(?:;\s*(?P<unread>\d+) unread messages?)?$")
GROUP = re.compile(r"^(?P<group>[^,]+), (?:Expanded|Collapsed)$")
LOAD_MORE = "Load more conversations"
TIME_TOKEN = re.compile(
    r"(?:,\s{2,}|\s{3,})"
    r"(?:(?P<status>Replied|Forwarded)\s+)?"
    r"(?P<time>\d{1,2}:\d{2}(?: [AP]M)?|Yesterday|"
    r"(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*|\d{1,2}/\d{1,2}/\d{4}),"
)
MEETING = re.compile(r",?\s*Meeting message(?:,|$)")
HAS_FILES = "Has files,"
UNREAD_COUNT = re.compile(r"^(\d+) unread messages?$")
MESSAGE_COUNT = re.compile(r"^(\d+) messages?$")
KNOWN_FLAGS = {
    "high priority",
    "low priority",
    "unread",
    "flagged",
    "completed",
    "has attachments",
    "has attachment",
    "replied",
    "forwarded",
    "draft",
    "pinned",
    "external",
    "encrypted",
    "signed",
    "mentioned",
    "expanded",
    "collapsed",
}
TIME = re.compile(r"^\d{1,2}:\d{2}(?:\s*[AP]M)?$", re.IGNORECASE)
DAY_MONTH_YEAR = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
WEEKDAYS = ("mon", "tue", "wed", "thu", "fri", "sat", "sun")
SPACES = re.compile(r"[\s​ ]+")


def clean(text: str) -> str:
    return SPACES.sub(" ", text).strip()


def parse_folder(description: str) -> Folder | None:
    match = FOLDER.match(clean(description))
    if not match or not match["name"]:
        return None
    return Folder(name=match["name"], unread=int(match["unread"] or 0))


def parse_rows(rows: list[Node], today: date) -> list[Row]:
    parsed: list[Row] = []
    group = ""
    children_left = 0
    for index, node in enumerate(rows):
        description = node.description
        header = GROUP.match(clean(description))
        if header:
            group = header["group"]
            continue
        if not description or clean(description) == LOAD_MORE:
            continue
        summary = parse_summary(description, index, group, today)
        if summary is None:
            continue
        if children_left > 0 and parsed:
            children_left -= 1
            parent = parsed[-1]
            thread = (*parent.summary.thread, thread_line(summary))
            parsed[-1] = Row(summary=replace(parent.summary, thread=thread), node=parent.node)
            continue
        if "expanded" in summary.flags:
            children_left = summary.message_count
        parsed.append(Row(summary=summary, node=node))
    return parsed


def thread_line(summary: Summary) -> str:
    return f"{summary.title} | {summary.time_label} | {summary.preview[:160]}"


def has_load_more(rows: list[Node]) -> Node | None:
    return next((node for node in rows if clean(node.description) == LOAD_MORE), None)


def parse_summary(description: str, index: int, group: str, today: date) -> Summary | None:
    raw = description.replace("\u202f", " ").replace("\u200b", "")
    match = TIME_TOKEN.search(raw)
    if not match:
        return None
    head = raw[: match.start()]
    rest = raw[match.end() :].strip()
    meeting = ""
    meeting_match = MEETING.search(head)
    if meeting_match:
        head, meeting = head[: meeting_match.start()], head[meeting_match.end() :]
    flags, unread_count, message_count, title = split_flags(head)
    if match["status"]:
        flags.append(match["status"].casefold())
    if rest.startswith(HAS_FILES):
        flags.append("has attachments")
        rest = rest.removeprefix(HAS_FILES)
    preview = clean(rest).rstrip(", ").strip()
    time_label = clean(match["time"])
    return Summary(
        key=fingerprint(title, time_label, preview),
        index=index,
        group=group,
        flags=tuple(flags),
        unread=unread_count > 0 or "unread" in flags,
        unread_count=unread_count or (1 if "unread" in flags else 0),
        message_count=message_count or 1,
        title=title,
        meeting=clean(meeting).strip(", "),
        time_label=time_label,
        date=label_date(time_label, group, today),
        preview=preview,
        thread=(),
    )


def split_flags(head: str) -> tuple[list[str], int, int, str]:
    parts = [clean(part) for part in head.split(",")]
    flags: list[str] = []
    unread_count = 0
    message_count = 0
    while parts and (not parts[0] or is_flag(parts[0])):
        part = parts.pop(0)
        if not part:
            continue
        unread = UNREAD_COUNT.match(part)
        messages = MESSAGE_COUNT.match(part)
        if unread:
            unread_count = int(unread[1])
        elif messages:
            message_count = int(messages[1])
        else:
            flags.append(part.casefold())
    return flags, unread_count, message_count, ", ".join(part for part in parts if part)


def is_flag(part: str) -> bool:
    lowered = part.casefold()
    return (
        lowered in KNOWN_FLAGS
        or UNREAD_COUNT.match(part) is not None
        or MESSAGE_COUNT.match(part) is not None
    )


def fingerprint(title: str, time_label: str, preview: str) -> str:
    digest = hashlib.sha1(f"{title}\n{time_label}\n{preview[:120]}".encode()).hexdigest()
    return digest[:12]


def label_date(label: str, group: str, today: date) -> str | None:
    found = parse_label(label, today)
    if found is None and group.casefold() == "today":
        found = today
    return found.isoformat() if found else None


def parse_label(label: str, today: date) -> date | None:
    text = label.strip()
    if TIME.match(text):
        return today
    lowered = text.casefold()
    if lowered == "yesterday":
        return today - timedelta(days=1)
    if lowered[:3] in WEEKDAYS and len(lowered.split()) == 1:
        days_back = (today.weekday() - WEEKDAYS.index(lowered[:3])) % 7 or 7
        return today - timedelta(days=days_back)
    match = DAY_MONTH_YEAR.match(text)
    if match:
        try:
            return date(int(match[3]), int(match[2]), int(match[1]))
        except ValueError:
            return None
    return None
