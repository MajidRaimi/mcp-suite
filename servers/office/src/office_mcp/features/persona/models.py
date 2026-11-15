from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any

from office_mcp.shared.drafts import DASH_CHARACTERS, NO_PERSONA_VERSION, Channel
from office_mcp.shared.recipients import normalize_email

DEFAULT_SIGN_OFF_KEY = "default"
DEFAULT_EMOJI_ALLOW = frozenset({"\u00a9", "\u00ae", "\u2122"})


class Language(StrEnum):
    EN = "en"
    AR = "ar"


class SignOffRule(StrEnum):
    REQUIRED = "required"
    FORBIDDEN = "forbidden"
    ANY = "any"


class CharacterSet(StrEnum):
    DASHES = "dashes"


CHARACTER_SETS: dict[CharacterSet, frozenset[str]] = {CharacterSet.DASHES: DASH_CHARACTERS}


class PersonaState(StrEnum):
    LOADED = "loaded"
    NONE = "none"
    MISSING = "missing"
    INVALID = "invalid"


@dataclass(frozen=True)
class LanguageRules:
    default: Language = Language.EN
    mirror: frozenset[Language] = frozenset()
    enforce: bool = True

    def to_record(self) -> dict[str, Any]:
        return {
            "default": self.default.value,
            "mirror": sorted(language.value for language in self.mirror),
            "enforce": self.enforce,
        }


@dataclass(frozen=True)
class ForbidRules:
    character_sets: tuple[CharacterSet, ...] = ()
    characters: tuple[str, ...] = ()
    emoji: bool = False
    emoji_allow: frozenset[str] = DEFAULT_EMOJI_ALLOW
    phrases: tuple[str, ...] = ()
    patterns: tuple[re.Pattern[str], ...] = ()

    def forbidden_characters(self) -> frozenset[str]:
        preset = (CHARACTER_SETS[name] for name in self.character_sets)
        return frozenset(self.characters).union(*preset)

    def to_record(self) -> dict[str, Any]:
        return {
            "character_sets": [name.value for name in self.character_sets],
            "characters": [code_point(char) for char in self.characters],
            "emoji": self.emoji,
            "emoji_allow": sorted(code_point(char) for char in self.emoji_allow),
            "phrases": list(self.phrases),
            "patterns": [pattern.pattern for pattern in self.patterns],
        }


@dataclass(frozen=True)
class ChannelRules:
    sign_off: Mapping[str, str] = field(default_factory=dict)
    sign_off_rule: SignOffRule = SignOffRule.ANY
    nothing_after_sign_off: bool = False
    applies_to_accounts: tuple[str, ...] = ()
    max_words: int | None = None
    max_chars: int | None = None

    def closing(self, language: Language) -> str | None:
        return self.sign_off.get(language.value) or self.sign_off.get(DEFAULT_SIGN_OFF_KEY)

    def closings(self) -> tuple[str, ...]:
        return tuple(dict.fromkeys(self.sign_off.values()))

    def applies_to(self, account: str) -> bool:
        if not self.applies_to_accounts or not account.strip():
            return True
        address = normalize_email(account)
        return any(account_matches(entry, address) for entry in self.applies_to_accounts)

    def to_record(self) -> dict[str, Any]:
        return {
            "sign_off": dict(self.sign_off),
            "sign_off_rule": self.sign_off_rule.value,
            "nothing_after_sign_off": self.nothing_after_sign_off,
            "applies_to_accounts": list(self.applies_to_accounts),
            "max_words": self.max_words,
            "max_chars": self.max_chars,
        }


@dataclass(frozen=True)
class Rules:
    version: int = 1
    language: LanguageRules | None = None
    forbid: ForbidRules = field(default_factory=ForbidRules)
    channels: Mapping[Channel, ChannelRules] = field(default_factory=dict)

    def channel(self, channel: Channel) -> ChannelRules | None:
        return self.channels.get(channel)

    def to_record(self) -> dict[str, Any]:
        return {
            "version": self.version,
            "language": None if self.language is None else self.language.to_record(),
            "forbid": self.forbid.to_record(),
            "channels": {
                channel.value: rules.to_record() for channel, rules in self.channels.items()
            },
        }


@dataclass(frozen=True)
class Persona:
    prose: str
    rules: Rules
    version: str = NO_PERSONA_VERSION
    source: Path | None = None
    modified_at: datetime | None = None
    loaded_at: datetime | None = None


@dataclass(frozen=True)
class PersonaStatus:
    state: PersonaState
    path: Path
    version: str
    modified_at: datetime | None
    permissions_ok: bool
    error: str | None = None

    def to_record(self) -> dict[str, Any]:
        return {
            "state": self.state.value,
            "path": str(self.path),
            "version": self.version,
            "modified_at": None if self.modified_at is None else self.modified_at.isoformat(),
            "permissions_ok": self.permissions_ok,
            "error": self.error,
        }


EMPTY_PERSONA = Persona(prose="", rules=Rules())


def account_matches(entry: str, address: str) -> bool:
    wanted = normalize_email(entry)
    if wanted.startswith("@"):
        return address.endswith(wanted)
    return address == wanted


def code_point(char: str) -> str:
    return f"U+{ord(char):04X}"
