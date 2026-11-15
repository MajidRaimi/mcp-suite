from __future__ import annotations

import os
from collections.abc import Sequence
from datetime import UTC, datetime
from pathlib import Path

from office_mcp.features.persona.bootstrap.sources import OwnMessage, SentMail
from office_mcp.features.persona.guard import PersonaGuard
from office_mcp.features.persona.store import PersonaStore
from office_mcp.shared.drafts import Channel, Draft, Review

LOADED_AT = datetime(2026, 10, 20, 9, 0, tzinfo=UTC)

FULL_FRONT_MATTER = """+++
version = 1

[language]
default = "en"
mirror = ["ar"]
enforce = true

[forbid]
character_sets = ["dashes"]
emoji = true
phrases = ["I hope this email finds you well", "أتمنى أن تكون بخير"]
patterns = ["(?<=\\\\S) - (?=\\\\S)", "--"]

[email]
sign_off = { default = "Best regards,", ar = "مع التحية" }
sign_off_rule = "required"
nothing_after_sign_off = true
applies_to_accounts = ["@example.com"]
max_words = 50

[teams]
sign_off_rule = "forbidden"
max_words = 20

[rsvp]
max_words = 10
+++
# Voice
Short and direct.
"""


def write_persona(path: Path, text: str, mode: int = 0o600) -> Path:
    path.write_text(text, encoding="utf-8")
    os.chmod(path, mode)
    return path


def persona_store(path: Path, explicit: bool = True, required: bool = False) -> PersonaStore:
    return PersonaStore(path=path, explicit=explicit, required=required, clock=lambda: LOADED_AT)


def guard_for(tmp_path: Path, text: str = FULL_FRONT_MATTER) -> PersonaGuard:
    return PersonaGuard(persona_store(write_persona(tmp_path / "persona.md", text)))


def review(
    guard: PersonaGuard,
    body: str,
    channel: Channel = Channel.EMAIL,
    context: str = "",
    context_author: str = "",
    account: str = "alex@example.com",
    subject: str = "",
    allow_language_mismatch: bool = False,
) -> Review:
    draft = Draft(
        channel=channel,
        body=body,
        subject=subject,
        context=context,
        context_author=context_author,
        sender_account=account,
        allow_language_mismatch=allow_language_mismatch,
    )
    return guard.review(draft)


def rules_hit(result: Review) -> list[str]:
    return [item.rule for item in result.violations]


class FakeSentMailSource:
    def __init__(self, mails: Sequence[SentMail]) -> None:
        self.mails = list(mails)
        self.limits: list[int] = []

    def sent_mail(self, limit: int) -> Sequence[SentMail]:
        self.limits.append(limit)
        return self.mails[:limit]


class FakeOwnChatSource:
    def __init__(self, messages: Sequence[OwnMessage]) -> None:
        self.messages = list(messages)
        self.calls: list[tuple[int, int]] = []

    def own_messages(self, chats: int, per_chat: int) -> Sequence[OwnMessage]:
        self.calls.append((chats, per_chat))
        return self.messages
