from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass

LIVE_FLAG = "OFFICE_MCP_LIVE"
SELF_THREAD_VARIABLE = "OFFICE_MCP_LIVE_SELF_THREAD"
RUN_VARIABLE = "OFFICE_MCP_LIVE_RUN"
SUBJECT_PREFIX = "[mcp-suite test]"
TEAMS_PREFIX = "teams:"


class LiveRefusal(Exception):
    pass


@dataclass(frozen=True)
class SelfTargets:
    account: str
    self_thread: str


def live_enabled(environ: Mapping[str, str]) -> bool:
    return environ.get(LIVE_FLAG, "").strip() == "1"


def text_field(status: Mapping[str, object], name: str) -> str:
    value = status.get(name)
    if not isinstance(value, str) or not value.strip():
        raise LiveRefusal(f"office_status has no {name}; the live run needs it")
    return value.strip()


def allowlist_entries(status: Mapping[str, object]) -> list[str]:
    entries = status.get("send_only_to")
    if not isinstance(entries, list):
        raise LiveRefusal("OFFICE_MCP_SEND_ONLY_TO is not set; the live run sends only to self")
    return sorted(str(entry).strip().lower() for entry in entries)


def self_targets(status: Mapping[str, object], environ: Mapping[str, str]) -> SelfTargets:
    account = text_field(status, "outlook_account").lower()
    thread = environ.get(SELF_THREAD_VARIABLE, "").strip() or text_field(status, "self_chat_thread")
    expected = sorted([account, f"{TEAMS_PREFIX}{thread}".lower()])
    found = allowlist_entries(status)
    if found != expected:
        raise LiveRefusal(f"send_only_to must be exactly {expected} for a live run, found {found}")
    return SelfTargets(account=account, self_thread=thread)


def run_subject(run_id: str, label: str) -> str:
    return f"{SUBJECT_PREFIX} {run_id} {label}"
