from __future__ import annotations

import re
from collections import Counter
from collections.abc import Iterable, Sequence
from dataclasses import dataclass

from office_mcp.features.outlook.compose.gateway import (
    RecipientField,
    RecipientRead,
    RecipientToken,
)
from office_mcp.features.outlook.conversation import LatestMessage
from office_mcp.features.outlook.models import Person
from office_mcp.shared.errors import ComposeError
from office_mcp.shared.recipients import is_email, normalize_email

MAX_PER_FIELD = 20
MAX_TOTAL = 40
EMBEDDED_EMAIL = re.compile(r"[A-Za-z0-9_%+'.-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")
TOKEN_FILLER = re.compile(r"[\s;,￼​⁠]+")


@dataclass(frozen=True)
class ExpectedRecipients:
    to: tuple[str, ...]
    cc: tuple[str, ...]
    bcc: tuple[str, ...]
    merge_to_cc: bool = False

    def everyone(self) -> tuple[str, ...]:
        return (*self.to, *self.cc, *self.bcc)


def validate_addresses(
    field: RecipientField, addresses: Sequence[str], minimum: int, maximum: int = MAX_PER_FIELD
) -> tuple[str, ...]:
    cleaned = tuple(address.strip() for address in addresses)
    if not minimum <= len(cleaned) <= maximum:
        raise ComposeError(f"{field.value} takes {minimum} to {maximum} addresses")
    for address in cleaned:
        if not is_email(address):
            raise ComposeError(
                f"{address!r} is not an email address; resolve it with outlook_people"
            )
    return cleaned


def require_distinct(addresses: Iterable[str]) -> None:
    counts = Counter(normalize_email(address) for address in addresses)
    repeated = sorted(address for address, count in counts.items() if count > 1)
    if repeated:
        raise ComposeError(f"each address may appear once across To, Cc and Bcc: {repeated}")
    if sum(counts.values()) > MAX_TOTAL:
        raise ComposeError(f"a message takes at most {MAX_TOTAL} recipients in total")


def person_address(person: Person | None, role: str) -> str:
    if person is None or not is_email(person.email):
        name = person.name if person is not None and person.name else "a recipient"
        raise ComposeError(f"could not read the {role} address of {name}; open the mail and check")
    return normalize_email(person.email)


def reply_recipients(latest: LatestMessage, owner: str | None, reply_all: bool) -> tuple[str, ...]:
    if reply_all:
        return reply_all_recipients(latest, owner)
    sender = person_address(latest.sender, "sender")
    if owner is None or sender != normalize_email(owner):
        return (sender,)
    original = tuple(person_address(person, "To") for person in latest.to)
    return original or (sender,)


def reply_all_recipients(latest: LatestMessage, owner: str | None) -> tuple[str, ...]:
    if owner is None:
        raise ComposeError(
            "reply all needs the owner's address; set OFFICE_MCP_OWNER_EMAIL or open the "
            "main Outlook window"
        )
    if latest.hidden_recipients:
        raise ComposeError(
            f"Outlook hides {latest.hidden_recipients} recipients of this mail, so reply all "
            "cannot be checked; reply to the sender and add the others as cc instead"
        )
    me = normalize_email(owner)
    people = [latest.sender, *latest.to, *latest.cc]
    addresses = [person_address(person, "recipient") for person in people]
    others = tuple(dict.fromkeys(address for address in addresses if address != me))
    return others or (me,)


def parse_token(label: str, extra: Sequence[str] = ()) -> RecipientToken:
    for text in (label, *extra):
        found = EMBEDDED_EMAIL.search(text)
        if found is not None and is_email(found.group(0)):
            return RecipientToken(label=label, address=normalize_email(found.group(0)))
    return RecipientToken(label=label, address=None)


def pending_text(value: str) -> str:
    return TOKEN_FILLER.sub(" ", value).strip()


def token_addresses(field: RecipientField, read: RecipientRead) -> tuple[str, ...]:
    if read.pending:
        raise ComposeError(
            f"the {field.value} field still holds text that is not a recipient: "
            f"{read.pending[:80]!r}"
        )
    unreadable = [token.label for token in read.tokens if token.address is None]
    if unreadable:
        raise ComposeError(
            f"could not read the address of a recipient in {field.value}: {unreadable[0][:80]!r}"
        )
    return tuple(token.address for token in read.tokens if token.address is not None)


def require_tokens(field: str, found: Sequence[str], expected: Sequence[str]) -> None:
    if len(found) != len(expected):
        raise ComposeError(
            f"{field} shows {len(found)} recipients but {len(expected)} were requested"
        )
    wanted = Counter(normalize_email(address) for address in expected)
    shown = Counter(normalize_email(address) for address in found)
    if wanted != shown:
        extra = sorted((shown - wanted).elements())
        missing = sorted((wanted - shown).elements())
        raise ComposeError(f"{field} recipients differ: unexpected {extra}, missing {missing}")
