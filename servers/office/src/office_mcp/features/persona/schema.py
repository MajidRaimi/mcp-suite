from __future__ import annotations

import re
from collections.abc import Mapping
from enum import StrEnum
from pathlib import Path

from office_mcp.features.persona.models import (
    DEFAULT_EMOJI_ALLOW,
    DEFAULT_SIGN_OFF_KEY,
    ChannelRules,
    CharacterSet,
    ForbidRules,
    Language,
    LanguageRules,
    Rules,
    SignOffRule,
)
from office_mcp.shared.drafts import Channel
from office_mcp.shared.errors import PersonaFormatError
from office_mcp.shared.recipients import DOMAIN_PATTERN, is_email

SUPPORTED_VERSION = 1
TOP_KEYS = frozenset({"version", "language", "forbid", *(channel.value for channel in Channel)})
LANGUAGE_KEYS = frozenset({"default", "mirror", "enforce"})
FORBID_KEYS = frozenset(
    {"character_sets", "characters", "emoji", "emoji_allow", "phrases", "patterns"}
)
CHANNEL_KEYS = frozenset(
    {
        "sign_off",
        "sign_off_rule",
        "nothing_after_sign_off",
        "applies_to_accounts",
        "max_words",
        "max_chars",
    }
)
SIGN_OFF_KEYS = frozenset({DEFAULT_SIGN_OFF_KEY, *(language.value for language in Language)})
TABLE_HEADER = re.compile(r"^\s*\[+\s*([^\]]+?)\s*\]+\s*$")
KEY_LINE = re.compile(r"^\s*([A-Za-z0-9_.\-\"' ]+?)\s*=")

type Table = Mapping[str, object]


class KeyPathError(PersonaFormatError):
    def __init__(self, key_path: str, problem: str) -> None:
        super().__init__(f"{key_path}: {problem}")
        self.key_path = key_path
        self.problem = problem


def parse_rules(data: Table, front_matter: str, path: Path) -> Rules:
    try:
        return build_rules(data)
    except KeyPathError as error:
        line = locate_key(front_matter, error.key_path)
        where = f"{path}: line {line}" if line is not None else f"{path}"
        raise PersonaFormatError(f"{where}: {error.key_path}: {error.problem}") from error


def build_rules(data: Table) -> Rules:
    require_known(data, "", TOP_KEYS)
    version = optional_int(data, "version", "", minimum=1)
    if version is not None and version != SUPPORTED_VERSION:
        raise KeyPathError("version", f"must be {SUPPORTED_VERSION}, got {version}")
    language = optional_table(data, "language", "")
    forbid = optional_table(data, "forbid", "")
    return Rules(
        version=SUPPORTED_VERSION,
        language=None if language is None else build_language(language),
        forbid=ForbidRules() if forbid is None else build_forbid(forbid),
        channels=build_channels(data),
    )


def build_language(data: Table) -> LanguageRules:
    require_known(data, "language", LANGUAGE_KEYS)
    default = data.get("default", Language.EN.value)
    mirror = string_list(data, "mirror", "language")
    return LanguageRules(
        default=choice(default, Language, "language.default"),
        mirror=frozenset(
            choice(item, Language, f"language.mirror[{index}]") for index, item in enumerate(mirror)
        ),
        enforce=boolean(data, "enforce", "language", default=True),
    )


def build_forbid(data: Table) -> ForbidRules:
    require_known(data, "forbid", FORBID_KEYS)
    sets = string_list(data, "character_sets", "forbid")
    allow = data.get("emoji_allow")
    return ForbidRules(
        character_sets=tuple(
            choice(item, CharacterSet, f"forbid.character_sets[{index}]")
            for index, item in enumerate(sets)
        ),
        characters=single_characters(data, "characters"),
        emoji=boolean(data, "emoji", "forbid", default=False),
        emoji_allow=(
            DEFAULT_EMOJI_ALLOW
            if allow is None
            else frozenset(single_characters(data, "emoji_allow"))
        ),
        phrases=non_empty_strings(data, "phrases", "forbid"),
        patterns=compile_patterns(data),
    )


def build_channels(data: Table) -> dict[Channel, ChannelRules]:
    channels: dict[Channel, ChannelRules] = {}
    for channel in Channel:
        table = optional_table(data, channel.value, "")
        if table is not None:
            channels[channel] = build_channel(table, channel.value)
    return channels


def build_channel(data: Table, prefix: str) -> ChannelRules:
    require_known(data, prefix, CHANNEL_KEYS)
    sign_off = build_sign_off(data.get("sign_off"), f"{prefix}.sign_off")
    rule = choice(
        data.get("sign_off_rule", SignOffRule.ANY.value), SignOffRule, f"{prefix}.sign_off_rule"
    )
    if rule is SignOffRule.REQUIRED and not sign_off:
        raise KeyPathError(f"{prefix}.sign_off", "is required when sign_off_rule is required")
    return ChannelRules(
        sign_off=sign_off,
        sign_off_rule=rule,
        nothing_after_sign_off=boolean(data, "nothing_after_sign_off", prefix, default=False),
        applies_to_accounts=accounts(data, prefix),
        max_words=optional_int(data, "max_words", prefix, minimum=1),
        max_chars=optional_int(data, "max_chars", prefix, minimum=1),
    )


def build_sign_off(value: object, key_path: str) -> dict[str, str]:
    if value is None:
        return {}
    if isinstance(value, str):
        return {DEFAULT_SIGN_OFF_KEY: required_text(value, key_path)}
    if not isinstance(value, Mapping):
        raise KeyPathError(key_path, f"must be a string or a table, got {type_name(value)}")
    require_known(value, key_path, SIGN_OFF_KEYS)
    return {str(key): required_text(item, f"{key_path}.{key}") for key, item in value.items()}


def accounts(data: Table, prefix: str) -> tuple[str, ...]:
    entries = string_list(data, "applies_to_accounts", prefix)
    for index, entry in enumerate(entries):
        if not is_account(entry):
            raise KeyPathError(
                f"{prefix}.applies_to_accounts[{index}]",
                f"must be an address or @domain, got {entry!r}",
            )
    return tuple(entry.strip().lower() for entry in entries)


def is_account(entry: str) -> bool:
    text = entry.strip()
    if text.startswith("@"):
        return DOMAIN_PATTERN.fullmatch(text[1:]) is not None
    return is_email(text)


def compile_patterns(data: Table) -> tuple[re.Pattern[str], ...]:
    compiled: list[re.Pattern[str]] = []
    for index, source in enumerate(non_empty_strings(data, "patterns", "forbid")):
        try:
            compiled.append(re.compile(source))
        except re.error as error:
            raise KeyPathError(
                f"forbid.patterns[{index}]", f"is not a valid regular expression: {error}"
            ) from error
    return tuple(compiled)


def single_characters(data: Table, key: str) -> tuple[str, ...]:
    items = string_list(data, key, "forbid")
    for index, item in enumerate(items):
        if len(item) != 1:
            raise KeyPathError(
                f"forbid.{key}[{index}]", f"must be exactly one character, got {item!r}"
            )
    return tuple(items)


def non_empty_strings(data: Table, key: str, prefix: str) -> tuple[str, ...]:
    items = string_list(data, key, prefix)
    for index, item in enumerate(items):
        required_text(item, f"{prefix}.{key}[{index}]")
    return tuple(items)


def require_known(data: Table, prefix: str, allowed: frozenset[str]) -> None:
    for key in sorted(set(data) - allowed):
        raise KeyPathError(join(prefix, key), "is not a known key")


def optional_table(data: Table, key: str, prefix: str) -> Table | None:
    value = data.get(key)
    if value is None:
        return None
    if not isinstance(value, Mapping):
        raise KeyPathError(join(prefix, key), f"must be a table, got {type_name(value)}")
    return value


def string_list(data: Table, key: str, prefix: str) -> list[str]:
    value = data.get(key, [])
    key_path = join(prefix, key)
    if not isinstance(value, list):
        raise KeyPathError(key_path, f"must be a list of strings, got {type_name(value)}")
    for index, item in enumerate(value):
        if not isinstance(item, str):
            raise KeyPathError(f"{key_path}[{index}]", f"must be a string, got {type_name(item)}")
    return [str(item) for item in value]


def boolean(data: Table, key: str, prefix: str, default: bool) -> bool:
    value = data.get(key, default)
    if not isinstance(value, bool):
        raise KeyPathError(join(prefix, key), f"must be true or false, got {type_name(value)}")
    return value


def optional_int(data: Table, key: str, prefix: str, minimum: int) -> int | None:
    value = data.get(key)
    if value is None:
        return None
    key_path = join(prefix, key)
    if isinstance(value, bool) or not isinstance(value, int):
        raise KeyPathError(key_path, f"must be a whole number, got {type_name(value)}")
    if value < minimum:
        raise KeyPathError(key_path, f"must be at least {minimum}, got {value}")
    return value


def required_text(value: object, key_path: str) -> str:
    if not isinstance(value, str):
        raise KeyPathError(key_path, f"must be a string, got {type_name(value)}")
    if not value.strip():
        raise KeyPathError(key_path, "must not be empty")
    return value.strip()


def choice[E: StrEnum](value: object, kind: type[E], key_path: str) -> E:
    allowed = [member.value for member in kind]
    if isinstance(value, str) and value in allowed:
        return kind(value)
    raise KeyPathError(key_path, f"must be one of {', '.join(allowed)}, got {value!r}")


def join(prefix: str, key: str) -> str:
    return f"{prefix}.{key}" if prefix else key


def type_name(value: object) -> str:
    names = {bool: "a boolean", int: "a number", float: "a number", str: "a string"}
    for kind, name in names.items():
        if type(value) is kind:
            return name
    if isinstance(value, list):
        return "a list"
    if isinstance(value, Mapping):
        return "a table"
    return type(value).__name__


def locate_key(front_matter: str, key_path: str) -> int | None:
    wanted = [part for part in re.sub(r"\[\d+\]", "", key_path).split(".") if part]
    best: tuple[int, int] | None = None
    table: list[str] = []
    for number, line in enumerate(front_matter.split("\n"), start=1):
        header = TABLE_HEADER.match(line)
        if header:
            table = split_key(header.group(1))
            candidate = table
        else:
            key = KEY_LINE.match(line)
            if key is None:
                continue
            candidate = [*table, *split_key(key.group(1))]
        depth = shared_prefix(candidate, wanted)
        if depth == len(candidate) and depth > 0 and (best is None or depth > best[0]):
            best = (depth, number)
    return None if best is None else best[1] + 1


def split_key(text: str) -> list[str]:
    return [part.strip().strip("\"'") for part in text.split(".") if part.strip()]


def shared_prefix(left: list[str], right: list[str]) -> int:
    depth = 0
    for first, second in zip(left, right, strict=False):
        if first != second:
            break
        depth += 1
    return depth
