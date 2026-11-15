from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any

from office_mcp.shared.node import Node


@dataclass(frozen=True)
class Folder:
    name: str
    unread: int
    total: int | None = None
    unseen: int | None = None
    level: int = 0
    path: str = ""
    favorite: bool = False

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class Summary:
    key: str
    index: int
    group: str
    flags: tuple[str, ...]
    unread: bool
    unread_count: int
    message_count: int
    title: str
    meeting: str
    time_label: str
    date: str | None
    preview: str
    thread: tuple[str, ...] = ()
    folder: str = ""
    view: str | None = None
    flagged: bool = False
    replied: bool = False
    forwarded: bool = False
    has_attachments: bool = False
    pinned: bool = False

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class Row:
    summary: Summary
    node: Node


@dataclass(frozen=True)
class Person:
    name: str
    email: str

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class Attachment:
    name: str
    size: str

    def to_record(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ThreadMessage:
    sender: str
    sent: str
    to: str
    cc: str
    subject: str
    body: str

    def to_record(self) -> dict[str, Any]:
        return {
            "from": self.sender,
            "sent": self.sent,
            "to": self.to,
            "cc": self.cc,
            "subject": self.subject,
            "body": self.body,
        }


@dataclass(frozen=True)
class Email:
    subject: str
    tags: tuple[str, ...]
    sender: Person | None
    recipients: dict[str, list[Person]]
    hidden_recipients: dict[str, int]
    sent: str
    notes: tuple[str, ...]
    body: str
    truncated: bool
    summary: Summary | None = None
    messages: tuple[ThreadMessage, ...] = ()
    attachments: tuple[Attachment, ...] = ()

    def to_record(self) -> dict[str, Any]:
        return {
            "subject": self.subject,
            "tags": list(self.tags),
            "from": self.sender.to_record() if self.sender else None,
            "recipients": {
                kind: [person.to_record() for person in people]
                for kind, people in self.recipients.items()
            },
            "hidden_recipients": self.hidden_recipients,
            "sent": self.sent,
            "notes": list(self.notes),
            "body": self.body,
            "truncated": self.truncated,
            "conversation": self.summary.to_record() if self.summary else None,
            "messages": [message.to_record() for message in self.messages],
            "attachments": [attachment.to_record() for attachment in self.attachments],
        }
