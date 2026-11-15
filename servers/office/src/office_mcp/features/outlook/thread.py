from __future__ import annotations

import re
from dataclasses import dataclass
from itertools import pairwise

from office_mcp.features.outlook.models import ThreadMessage

HEADER_LINE = re.compile(r"^(?P<label>From|Sent|Date|To|Cc|Subject):\s*(?P<value>.*)$")
LABEL_ONLY = re.compile(r"^(?:From|Sent|Date|To|Cc|Subject):\s*$")
FROM_LABEL = "From"
DATE_LABELS = frozenset({"Sent", "Date"})
LOOKAHEAD = 3
DEFAULT_LIMIT = 20


@dataclass(frozen=True)
class Block:
    header: dict[str, str]
    body: tuple[str, ...]


def split_thread(body: str, limit: int = DEFAULT_LIMIT) -> tuple[ThreadMessage, ...]:
    lines = joined_lines(body)
    starts = header_starts(lines)
    bounds = pairwise([*starts, len(lines)])
    messages = [to_message(read_block(lines[start:end])) for start, end in bounds]
    return tuple(messages[: max(limit, 0)])


def top_text(body: str) -> str:
    lines = joined_lines(body)
    starts = header_starts(lines)
    top = lines[: starts[0]] if starts else lines
    return "\n".join(top).strip()


def joined_lines(body: str) -> list[str]:
    raw = body.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines: list[str] = []
    index = 0
    while index < len(raw):
        line = raw[index].strip()
        if LABEL_ONLY.match(line) and index + 1 < len(raw):
            line = f"{line} {raw[index + 1].strip()}"
            index += 1
        lines.append(line)
        index += 1
    return lines


def header_starts(lines: list[str]) -> list[int]:
    return [index for index in range(len(lines)) if starts_header(lines, index)]


def starts_header(lines: list[str], index: int) -> bool:
    if label_of(lines[index]) != FROM_LABEL:
        return False
    following = lines[index + 1 : index + 1 + LOOKAHEAD]
    return any(label_of(line) in DATE_LABELS for line in following)


def label_of(line: str) -> str:
    match = HEADER_LINE.match(line)
    return match["label"] if match else ""


def read_block(lines: list[str]) -> Block:
    header: dict[str, str] = {}
    index = 0
    while index < len(lines):
        match = HEADER_LINE.match(lines[index])
        if match is None or match["label"] in header:
            break
        header[match["label"]] = match["value"].strip()
        index += 1
    return Block(header=header, body=tuple(lines[index:]))


def to_message(block: Block) -> ThreadMessage:
    header = block.header
    return ThreadMessage(
        sender=header.get("From", ""),
        sent=header.get("Sent") or header.get("Date", ""),
        to=header.get("To", ""),
        cc=header.get("Cc", ""),
        subject=header.get("Subject", ""),
        body="\n".join(block.body).strip(),
    )
