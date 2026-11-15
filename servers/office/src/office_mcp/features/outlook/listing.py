from __future__ import annotations

import hashlib
import re
from dataclasses import replace
from datetime import date, timedelta

from office_mcp.features.outlook.models import Folder, Row, Summary
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize

FOLDER = re.compile(
    r"^(?P<name>.+?)(?:;\s*(?P<count>\d+)\s+(?:(?P<kind>unread|unseen)\s+)?messages?)?$"
)
FAVORITES = "favorites"
PATH_SEPARATOR = "/"
FOLDER_TOKEN = re.compile(r"^\s*Folder:\s*(?P<folder>[^,]+),")
PINNED_GROUP = "pinned"
EXPANDED = "expanded"
ATTACHMENT_FLAGS = frozenset({"has attachments", "has attachment"})
GROUP = re.compile(r"^(?P<group>[^,]+), (?:Expanded|Collapsed)$")
LOAD_MORE = "Load more conversations"
TIME_TOKEN = re.compile(
    r"(?:,\s{2,}|\s{3,})"
    r"(?:(?P<status>Replied|Forwarded)\s+)?"
    r"(?P<time>\d{1,2}:\d{2}(?: [AP]M)?|Yesterday|"
    r"(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*|\d{1,2}/\d{1,2}/\d{4}),"
)
MEETING = re.compile(r",?\s*Meeting message(?:,|$)")
STATUS_TOKENS = {
    "has files": "has attachments",
    "has reactions": "has reactions",
    "flagged": "flagged",
    "completed": "completed",
    "marked as completed": "completed",
}
STATUS_TOKEN = re.compile(
    r"^\s*(?P<status>" + "|".join(re.escape(token) for token in STATUS_TOKENS) + r")(?:,|\s{2,}|$)",
    re.IGNORECASE,
)
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


def parse_folder(description: str, level: int = 0) -> Folder | None:
    match = FOLDER.match(clean(description))
    if not match or not match["name"]:
        return None
    count = int(match["count"]) if match["count"] else None
    kind = match["kind"]
    return Folder(
        name=match["name"],
        unread=count if kind == "unread" and count is not None else 0,
        total=count if kind is None else None,
        unseen=count if kind == "unseen" else None,
        level=level,
        path=match["name"],
    )


def parse_folders(nodes: list[Node]) -> list[Folder]:
    return [folder for _, folder in folder_pairs(nodes)]


def folder_pairs(nodes: list[Node]) -> list[tuple[Node, Folder]]:
    pairs: list[tuple[Node, Folder]] = []
    ancestors: list[str] = []
    for node in nodes:
        level = max(node.level or 0, 0)
        folder = parse_folder(node.description, level)
        if folder is None:
            continue
        ancestors = [*ancestors[:level], folder.name]
        favorite = level > 0 and normalize(ancestors[0]) == FAVORITES
        located = replace(folder, path=PATH_SEPARATOR.join(ancestors), favorite=favorite)
        pairs.append((node, located))
    return pairs


def folder_matches(folder: Folder, wanted: str) -> bool:
    target = normalize(wanted)
    return normalize(folder.name) == target or normalize(folder.path) == target


def pick_folder_row(nodes: list[Node], wanted: str) -> Node | None:
    pairs = folder_pairs(nodes)
    target = normalize(wanted)
    exact = [pair for pair in pairs if folder_matches(pair[1], wanted)]
    prefixed = [pair for pair in pairs if normalize(pair[1].name).startswith(target)]
    for candidates in (exact, prefixed):
        preferred = [pair for pair in candidates if not pair[1].favorite] or candidates
        if preferred:
            return preferred[0][0]
    return None


def row_flags(flags: tuple[str, ...]) -> dict[str, bool]:
    present = set(flags)
    return {
        "flagged": "flagged" in present,
        "replied": "replied" in present,
        "forwarded": "forwarded" in present,
        "has_attachments": bool(present & ATTACHMENT_FLAGS),
        "pinned": "pinned" in present,
    }


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
            parsed[-1] = Row(summary=with_child(parent.summary, summary), node=parent.node)
            continue
        if "expanded" in summary.flags:
            children_left = summary.message_count
        parsed.append(Row(summary=summary, node=node))
    return parsed


def with_child(parent: Summary, child: Summary) -> Summary:
    thread = (*parent.thread, thread_line(child))
    if parent.preview or not child.preview:
        return replace(parent, thread=thread)
    key = fingerprint(parent.title, parent.time_label, child.preview)
    return replace(parent, thread=thread, preview=child.preview, key=key)


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
    folder = ""
    folder_match = FOLDER_TOKEN.match(rest)
    if folder_match:
        folder = clean(folder_match["folder"])
        rest = rest[folder_match.end() :].strip()
    status, rest = split_status(rest)
    flags.extend(status)
    preview = clean(rest).rstrip(", ").strip()
    time_label = clean(match["time"])
    booleans = row_flags(tuple(flags))
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
        folder=folder,
        flagged=booleans["flagged"],
        replied=booleans["replied"],
        forwarded=booleans["forwarded"],
        has_attachments=booleans["has_attachments"],
        pinned=booleans["pinned"] or group.casefold() == PINNED_GROUP,
    )


def split_status(rest: str) -> tuple[list[str], str]:
    found: list[str] = []
    while (match := STATUS_TOKEN.match(rest)) is not None:
        found.append(STATUS_TOKENS[clean(match["status"]).casefold()])
        rest = rest[match.end() :].strip()
    return found, rest


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


def expanded_header_of(header: Summary, original: Summary) -> bool:
    if EXPANDED not in header.flags or EXPANDED in original.flags:
        return False
    if header.time_label != original.time_label or not header.title:
        return False
    return original.title == header.title or original.title.endswith(f", {header.title}")


def fingerprint(title: str, time_label: str, preview: str) -> str:
    subject_tail = title.rsplit(", ", 1)[-1]
    digest = hashlib.sha1(f"{subject_tail}\n{time_label}\n{preview[:120]}".encode()).hexdigest()
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
