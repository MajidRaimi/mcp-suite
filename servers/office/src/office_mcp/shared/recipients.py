from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass

from office_mcp.shared.errors import RecipientNotAllowedError, SettingsError

SELF_ENTRY = "teams:self"
TEAMS_PREFIX = "teams:"
ALLOWLIST_VARIABLE = "OFFICE_MCP_SEND_ONLY_TO"
DOMAIN = r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)+"
EMAIL_PATTERN = re.compile(rf"[a-z0-9_%+'-]+(?:\.[a-z0-9_%+'-]+)*@{DOMAIN}", re.IGNORECASE)
DOMAIN_PATTERN = re.compile(DOMAIN, re.IGNORECASE)
THREAD_PATTERN = re.compile(r"[^\s,]+")


@dataclass(frozen=True)
class Allowlist:
    emails: frozenset[str]
    domains: frozenset[str]
    teams_threads: frozenset[str]
    teams_self: bool

    def entries(self) -> tuple[str, ...]:
        return (
            *sorted(self.emails),
            *(f"@{domain}" for domain in sorted(self.domains)),
            *((SELF_ENTRY,) if self.teams_self else ()),
            *(f"{TEAMS_PREFIX}{thread}" for thread in sorted(self.teams_threads)),
        )

    def allows_email(self, address: str) -> bool:
        normalized = normalize_email(address)
        if not is_email(normalized):
            return False
        return normalized in self.emails or normalized.rpartition("@")[2] in self.domains

    def allows_thread(self, thread_id: str, self_thread: str) -> bool:
        if thread_id in self.teams_threads:
            return True
        return self.teams_self and bool(self_thread) and thread_id == self_thread


def is_email(text: str) -> bool:
    return EMAIL_PATTERN.fullmatch(text.strip()) is not None


def normalize_email(text: str) -> str:
    return text.strip().lower()


def parse_allowlist(raw: str | None) -> Allowlist | None:
    if raw is None or not raw.strip():
        return None
    entries = [entry.strip() for entry in raw.split(",") if entry.strip()]
    if not entries:
        raise SettingsError(f"{ALLOWLIST_VARIABLE} is set but lists no entries")
    emails: set[str] = set()
    domains: set[str] = set()
    threads: set[str] = set()
    teams_self = False
    for entry in entries:
        if entry.lower() == SELF_ENTRY:
            teams_self = True
        elif entry.lower().startswith(TEAMS_PREFIX):
            threads.add(parse_thread_entry(entry))
        elif entry.startswith("@"):
            domains.add(parse_domain_entry(entry))
        elif is_email(entry):
            emails.add(normalize_email(entry))
        else:
            raise SettingsError(invalid_entry(entry))
    return Allowlist(frozenset(emails), frozenset(domains), frozenset(threads), teams_self)


def parse_thread_entry(entry: str) -> str:
    thread = entry[len(TEAMS_PREFIX) :].strip()
    if not THREAD_PATTERN.fullmatch(thread):
        raise SettingsError(invalid_entry(entry))
    return thread


def parse_domain_entry(entry: str) -> str:
    domain = entry[1:].strip().lower()
    if not DOMAIN_PATTERN.fullmatch(domain):
        raise SettingsError(invalid_entry(entry))
    return domain


def invalid_entry(entry: str) -> str:
    return (
        f"{ALLOWLIST_VARIABLE} entry {entry!r} must be user@domain, @domain, "
        f"{SELF_ENTRY} or {TEAMS_PREFIX}<thread id>"
    )


class RecipientGuard:
    def __init__(self, allowlist: Allowlist | None, self_thread: str) -> None:
        self._allowlist = allowlist
        self._self_thread = self_thread.strip()

    @property
    def active(self) -> bool:
        return self._allowlist is not None

    def require_emails(self, addresses: Sequence[str]) -> None:
        if self._allowlist is None:
            return
        blocked = [item for item in addresses if not self._allowlist.allows_email(item)]
        if blocked:
            raise RecipientNotAllowedError(self._refusal(blocked))

    def require_thread(self, thread_id: str) -> None:
        if self._allowlist is None:
            return
        if not self._allowlist.allows_thread(thread_id, self._self_thread):
            raise RecipientNotAllowedError(self._refusal([f"{TEAMS_PREFIX}{thread_id}"]))

    def describe(self) -> list[str] | None:
        return None if self._allowlist is None else list(self._allowlist.entries())

    def _refusal(self, blocked: Sequence[str]) -> str:
        allowed = ", ".join(self.describe() or [])
        return (
            f"Not sent: {', '.join(blocked)} not allowed by {ALLOWLIST_VARIABLE} "
            f"(allowed: {allowed}). Nothing was sent."
        )
