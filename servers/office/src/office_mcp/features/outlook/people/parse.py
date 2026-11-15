from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass
from enum import StrEnum
from typing import Any

from office_mcp.features.outlook.listing import clean
from office_mcp.features.outlook.models import Person
from office_mcp.shared.text import normalize

PEOPLE_SUGGESTION = re.compile(
    r"people suggestion\s+\d+\s+of\s+\d+(?:\s+highlighted)?\.?\s*", re.IGNORECASE
)
TYPED_MARK = "best match"
EMAIL = re.compile(r"[\w.+'-]+@[\w-]+(?:\.[\w-]+)+")
NAME_TRIM = " ,;:<>()[]"


class ContactSource(StrEnum):
    SUGGESTION = "outlook_suggestion"
    RECENT = "recent_recipient"


@dataclass(frozen=True)
class Contact:
    name: str
    email: str
    source: ContactSource

    def to_record(self) -> dict[str, Any]:
        return {"name": self.name, "email": self.email, "source": self.source.value}


def parse_suggestion(text: str) -> Contact | None:
    line = clean(text)
    marker = PEOPLE_SUGGESTION.search(line)
    if marker is None:
        return None
    rest = line[marker.end() :]
    emails = EMAIL.findall(rest)
    if not emails:
        return None
    email = emails[-1]
    name = clean(rest.replace(email, " ")).strip(NAME_TRIM)
    return Contact(name=name or email, email=email, source=ContactSource.SUGGESTION)


def parse_suggestions(texts: Iterable[str]) -> list[Contact]:
    return unique_contacts(
        contact for contact in (parse_suggestion(text) for text in texts) if contact is not None
    )


def typed_rows(texts: list[str]) -> list[str] | None:
    typed = [text for text in texts if TYPED_MARK in text.casefold()]
    return typed or None


def unique_contacts(contacts: Iterable[Contact]) -> list[Contact]:
    seen: set[str] = set()
    unique: list[Contact] = []
    for contact in contacts:
        address = contact.email.casefold()
        if address not in seen:
            seen.add(address)
            unique.append(contact)
    return unique


def contact_from(person: Person) -> Contact | None:
    if not EMAIL.fullmatch(person.email):
        return None
    return Contact(
        name=person.name or person.email, email=person.email, source=ContactSource.RECENT
    )


def matches_query(contact: Contact, query: str) -> bool:
    wanted = normalize(query)
    return bool(wanted) and (
        wanted in normalize(contact.name) or wanted in contact.email.casefold()
    )
