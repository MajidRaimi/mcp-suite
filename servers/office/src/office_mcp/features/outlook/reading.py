from __future__ import annotations

import re

from office_mcp.features.outlook.listing import clean
from office_mcp.features.outlook.models import Email, Person
from office_mcp.shared.node import Node

HEADER_DETAILS = "Message Header Details"
READING_PANE = "Reading Pane"
WEB_AREA = "AXWebArea"
STATIC_TEXT = "AXStaticText"
BUTTON = "AXButton"
FROM_PREFIX = "From:"
FIELD_LABEL = re.compile(r"^(To|Cc|Bcc|Required|Optional|Resources):$", re.IGNORECASE)
MORE = re.compile(r"^\+\s*(\d+)\s+more$")
EMAIL = re.compile(r"[\w.+'-]+@[\w-]+(?:\.[\w-]+)+")
NOTE_HINTS = (
    "importance",
    "you don't often get email",
    "this message was sent",
    "you replied",
    "you forwarded",
)
JOIN_TIGHT_BEFORE = (">", ")", ".", ",", ":", ";", "|", "!", "?")
JOIN_TIGHT_AFTER = ("<", "(", "/", "@")
DEFAULT_MAX_BODY = 20000


def parse_email(pane: Node, max_body: int = DEFAULT_MAX_BODY) -> Email | None:
    header = pane.find(lambda node: node.description == HEADER_DETAILS)
    if header is None:
        return None
    holder = next((child for child in pane.children if child.find(lambda n: n is header)), pane)
    subject, tags = subject_and_tags(holder, header)
    sender, recipients, hidden, sent, notes = parse_header(header)
    web = pane.find(lambda node: node.role == WEB_AREA)
    body = body_text(web) if web else ""
    return Email(
        subject=subject,
        tags=tags,
        sender=sender,
        recipients=recipients,
        hidden_recipients=hidden,
        sent=sent,
        notes=notes,
        body=body[:max_body],
        truncated=len(body) > max_body,
    )


def subject_and_tags(pane: Node, header: Node) -> tuple[str, tuple[str, ...]]:
    texts: list[str] = []
    for node in pane.children:
        if node is header or node.role == BUTTON:
            continue
        if node.find(lambda child: child is header) is not None:
            continue
        value = clean(node.value or node.title)
        if value and node.role in (STATIC_TEXT, "AXUnknown", "", "AXGroup"):
            texts.append(value)
        texts.extend(
            clean(child.value)
            for child in node.children
            if child.value and child.role != BUTTON and clean(child.value) not in texts
        )
    if not texts:
        return "", ()
    return texts[0], tuple(dict.fromkeys(texts[1:]))


def parse_header(
    header: Node,
) -> tuple[Person | None, dict[str, list[Person]], dict[str, int], str, tuple[str, ...]]:
    sender: Person | None = None
    recipients: dict[str, list[Person]] = {}
    hidden: dict[str, int] = {}
    sent = ""
    notes: list[str] = []
    field = "to"
    for node in header.walk():
        value = clean(node.value)
        if not value or node is header:
            continue
        label = FIELD_LABEL.match(value)
        if label:
            field = label[1].casefold()
        elif value.startswith(FROM_PREFIX) and node.role != BUTTON:
            sender = parse_person(value.removeprefix(FROM_PREFIX))
        elif node.role == BUTTON and sender is not None and EMAIL.search(value):
            person = parse_person(value)
            if person and person != sender and person not in recipients.get(field, []):
                recipients.setdefault(field, []).append(person)
        elif node.role == BUTTON and (more := MORE.match(value)):
            hidden[field] = int(more[1])
        elif node.role == STATIC_TEXT and any(hint in value.casefold() for hint in NOTE_HINTS):
            notes.append(value)
        elif node.role == STATIC_TEXT and not sent and looks_like_date(value):
            sent = value
    return sender, recipients, hidden, sent, tuple(notes)


def parse_person(text: str) -> Person | None:
    parts = [clean(part) for part in text.split(",")]
    email = next((part for part in parts if EMAIL.fullmatch(part)), "")
    name = next((part for part in parts if part and part != email), "")
    if not name and not email:
        return None
    return Person(name=name or email, email=email)


def looks_like_date(value: str) -> bool:
    lowered = value.casefold()
    return bool(re.search(r"\d{1,2}:\d{2}", value)) and (
        " at " in lowered or "/" in value or "am" in lowered or "pm" in lowered
    )


def body_text(web: Node) -> str:
    pieces = [clean(node.value) for node in web.walk() if node.role == STATIC_TEXT]
    text = ""
    for piece in pieces:
        if not piece:
            continue
        if not text:
            text = piece
        elif piece.startswith(JOIN_TIGHT_BEFORE) or text.endswith(JOIN_TIGHT_AFTER):
            text += piece
        else:
            text += "\n" + piece
    return text
