from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import UTC, datetime, tzinfo
from pathlib import Path
from typing import Literal
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from office_mcp.shared.errors import SettingsError
from office_mcp.shared.recipients import Allowlist, is_email, parse_allowlist
from office_mcp.shared.registry import ToolGroup, ToolPolicy

Transport = Literal["stdio", "streamable-http"]

PREFIX = "OFFICE_MCP_"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8769
DEFAULT_TRANSPORT: Transport = "stdio"
TRANSPORTS: tuple[Transport, ...] = ("stdio", "streamable-http")
MAX_PORT = 65535
LOCALTIME_PATH = "/etc/localtime"
ZONEINFO_MARKER = "zoneinfo/"
ALL = "all"
TRUE_WORDS = frozenset({"1", "true", "yes", "on"})
FALSE_WORDS = frozenset({"0", "false", "no", "off"})
DEFAULT_PERSONA_FILE = Path("~/.config/office-mcp/persona.md")
DEFAULT_LOCK_FILE = Path("~/Library/Caches/office-mcp/ui.lock")
DEFAULT_SELF_CHAT_THREAD = "48:notes"
DEFAULT_ACTION_TIMEOUT = 12.0
DEFAULT_IDLE_SECONDS = 1.5
DEFAULT_LOCK_TIMEOUT = 20.0
DEFAULT_STALE_AFTER = 10.0
DEFAULT_TEAMS_MAX_MEMBERS = 25
MIN_TEAMS_MEMBERS = 2
MAX_TEAMS_MEMBERS = 1000


@dataclass(frozen=True)
class Zone:
    name: str
    info: tzinfo


@dataclass(frozen=True)
class Settings:
    zone: Zone
    host: str
    port: int
    transport: Transport
    token: str | None
    allowed_hosts: tuple[str, ...]
    owner_name: str | None
    tool_policy: ToolPolicy
    persona_file: Path
    persona_explicit: bool
    persona_required: bool
    send_only_to: Allowlist | None
    self_chat_thread: str
    restore_frontmost: bool
    action_timeout: float
    idle_seconds: float
    lock_file: Path
    lock_timeout: float
    stale_after: float
    download_dir: Path | None
    teams_max_members: int
    owner_email: str | None

    @classmethod
    def from_env(cls, environ: Mapping[str, str] | None = None) -> Settings:
        env = os.environ if environ is None else environ
        return cls(
            zone=resolve_zone(env_value(env, "TIMEZONE"), env),
            host=env_value(env, "HOST") or DEFAULT_HOST,
            port=parse_port(env_value(env, "PORT"), f"{PREFIX}PORT"),
            transport=parse_transport(env_value(env, "TRANSPORT"), f"{PREFIX}TRANSPORT"),
            token=env_value(env, "TOKEN"),
            allowed_hosts=parse_allowed_hosts(env_value(env, "ALLOWED_HOSTS")),
            owner_name=env_value(env, "OWNER_NAME"),
            tool_policy=parse_tool_policy(env),
            persona_file=parse_path(env_value(env, "PERSONA_FILE"), DEFAULT_PERSONA_FILE),
            persona_explicit=env_value(env, "PERSONA_FILE") is not None,
            persona_required=parse_bool(
                env_value(env, "PERSONA_REQUIRED"), f"{PREFIX}PERSONA_REQUIRED", False
            ),
            send_only_to=parse_allowlist(env_value(env, "SEND_ONLY_TO")),
            self_chat_thread=env_value(env, "SELF_CHAT_THREAD") or DEFAULT_SELF_CHAT_THREAD,
            restore_frontmost=parse_bool(
                env_value(env, "RESTORE_FRONTMOST"), f"{PREFIX}RESTORE_FRONTMOST", True
            ),
            action_timeout=parse_seconds(
                env_value(env, "ACTION_TIMEOUT"),
                f"{PREFIX}ACTION_TIMEOUT",
                DEFAULT_ACTION_TIMEOUT,
                minimum=1,
                maximum=120,
            ),
            idle_seconds=parse_seconds(
                env_value(env, "IDLE_SECONDS"),
                f"{PREFIX}IDLE_SECONDS",
                DEFAULT_IDLE_SECONDS,
                minimum=0,
                maximum=30,
            ),
            lock_file=parse_path(env_value(env, "LOCK_FILE"), DEFAULT_LOCK_FILE),
            lock_timeout=parse_seconds(
                env_value(env, "LOCK_TIMEOUT"),
                f"{PREFIX}LOCK_TIMEOUT",
                DEFAULT_LOCK_TIMEOUT,
                minimum=1,
                maximum=300,
            ),
            stale_after=parse_seconds(
                env_value(env, "STALE_AFTER"),
                f"{PREFIX}STALE_AFTER",
                DEFAULT_STALE_AFTER,
                minimum=1,
                maximum=300,
            ),
            download_dir=parse_download_dir(env_value(env, "DOWNLOAD_DIR")),
            teams_max_members=parse_int(
                env_value(env, "TEAMS_MAX_MEMBERS"),
                f"{PREFIX}TEAMS_MAX_MEMBERS",
                DEFAULT_TEAMS_MAX_MEMBERS,
                minimum=MIN_TEAMS_MEMBERS,
                maximum=MAX_TEAMS_MEMBERS,
            ),
            owner_email=parse_owner_email(env_value(env, "OWNER_EMAIL")),
        )


def env_value(env: Mapping[str, str], suffix: str) -> str | None:
    return env.get(f"{PREFIX}{suffix}", "").strip() or None


def resolve_zone(name: str | None, env: Mapping[str, str], localtime: str = LOCALTIME_PATH) -> Zone:
    if name is None:
        return local_zone(env, localtime)
    zone = named_zone(name)
    if zone is None:
        raise SettingsError(f"{PREFIX}TIMEZONE {name!r} is not a known IANA zone")
    return zone


def local_zone(env: Mapping[str, str], localtime: str = LOCALTIME_PATH) -> Zone:
    tz_variable = env.get("TZ", "").strip().removeprefix(":")
    for name in (tz_variable, localtime_link_name(localtime)):
        zone = named_zone(name)
        if zone is not None:
            return zone
    fixed = datetime.now().astimezone().tzinfo
    if fixed is None:
        return Zone("UTC", UTC)
    return Zone(fixed.tzname(None) or "UTC", fixed)


def localtime_link_name(localtime: str) -> str:
    try:
        target = os.path.realpath(localtime, strict=True)
    except OSError:
        return ""
    _, marker, name = target.partition(ZONEINFO_MARKER)
    return name if marker else ""


def named_zone(name: str) -> Zone | None:
    if not name:
        return None
    try:
        return Zone(name, ZoneInfo(name))
    except (ZoneInfoNotFoundError, ValueError):
        return None


def parse_port(raw: str | None, name: str, default: int = DEFAULT_PORT) -> int:
    if raw is None or not raw.strip():
        return default
    try:
        port = int(raw.strip())
    except ValueError as error:
        raise SettingsError(f"{name} must be a whole number, got {raw!r}") from error
    if not 1 <= port <= MAX_PORT:
        raise SettingsError(f"{name} must be between 1 and {MAX_PORT}, got {port}")
    return port


def parse_transport(raw: str | None, name: str) -> Transport:
    if raw is None or not raw.strip():
        return DEFAULT_TRANSPORT
    value = raw.strip().lower()
    for transport in TRANSPORTS:
        if value == transport:
            return transport
    raise SettingsError(f"{name} must be one of {', '.join(TRANSPORTS)}, got {raw!r}")


def parse_allowed_hosts(raw: str | None) -> tuple[str, ...]:
    if raw is None:
        return ()
    hosts = tuple(entry.strip() for entry in raw.split(",") if entry.strip())
    for host in hosts:
        if "/" in host or any(character.isspace() for character in host):
            raise SettingsError(
                f"{PREFIX}ALLOWED_HOSTS entries must look like host or host:port, got {host!r}"
            )
    return hosts


def parse_tool_policy(env: Mapping[str, str]) -> ToolPolicy:
    disabled = parse_names(env_value(env, "DISABLED_TOOLS"))
    if parse_bool(env_value(env, "READ_ONLY"), f"{PREFIX}READ_ONLY", False):
        return ToolPolicy(
            groups=frozenset({ToolGroup.READ}),
            disabled=disabled,
            experimental=frozenset(),
            all_experimental=False,
        )
    experimental = parse_names(env_value(env, "EXPERIMENTAL"))
    return ToolPolicy(
        groups=parse_groups(env_value(env, "TOOLS")),
        disabled=disabled,
        experimental=frozenset(name for name in experimental if name.lower() != ALL),
        all_experimental=any(name.lower() == ALL for name in experimental),
    )


def parse_groups(raw: str | None) -> frozenset[ToolGroup]:
    names = {name.lower() for name in parse_names(raw)}
    if ALL in names:
        return frozenset(ToolGroup)
    groups = {ToolGroup.READ}
    known = {group.value: group for group in ToolGroup}
    for name in sorted(names):
        if name not in known:
            choices = ", ".join([*known, ALL])
            raise SettingsError(f"{PREFIX}TOOLS group {name!r} is not one of {choices}")
        groups.add(known[name])
    return frozenset(groups)


def parse_names(raw: str | None) -> frozenset[str]:
    if raw is None:
        return frozenset()
    return frozenset(entry.strip() for entry in raw.split(",") if entry.strip())


def parse_bool(raw: str | None, name: str, default: bool) -> bool:
    if raw is None or not raw.strip():
        return default
    value = raw.strip().lower()
    if value in TRUE_WORDS:
        return True
    if value in FALSE_WORDS:
        return False
    raise SettingsError(f"{name} must be true or false, got {raw!r}")


def parse_path(raw: str | None, default: Path) -> Path:
    if raw is None or not raw.strip():
        return default.expanduser()
    return Path(raw.strip()).expanduser()


def parse_download_dir(raw: str | None) -> Path | None:
    if raw is None or not raw.strip():
        return None
    path = Path(raw.strip()).expanduser()
    if not path.is_absolute():
        raise SettingsError(f"{PREFIX}DOWNLOAD_DIR must be an absolute path, got {raw!r}")
    return path


def parse_seconds(
    raw: str | None, name: str, default: float, minimum: float, maximum: float
) -> float:
    if raw is None or not raw.strip():
        return default
    try:
        seconds = float(raw.strip())
    except ValueError as error:
        raise SettingsError(f"{name} must be a number of seconds, got {raw!r}") from error
    if not minimum <= seconds <= maximum:
        raise SettingsError(f"{name} must be between {minimum:g} and {maximum:g}, got {raw!r}")
    return seconds


def parse_int(raw: str | None, name: str, default: int, minimum: int, maximum: int) -> int:
    if raw is None or not raw.strip():
        return default
    try:
        number = int(raw.strip())
    except ValueError as error:
        raise SettingsError(f"{name} must be a whole number, got {raw!r}") from error
    if not minimum <= number <= maximum:
        raise SettingsError(f"{name} must be between {minimum} and {maximum}, got {number}")
    return number


def parse_owner_email(raw: str | None) -> str | None:
    if raw is None or not raw.strip():
        return None
    if not is_email(raw):
        raise SettingsError(f"{PREFIX}OWNER_EMAIL must be an email address, got {raw!r}")
    return raw.strip().lower()
