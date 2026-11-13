from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import UTC, datetime, tzinfo
from typing import Literal
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from office_mcp.shared.errors import SettingsError

Transport = Literal["stdio", "streamable-http"]

PREFIX = "OFFICE_MCP_"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8769
DEFAULT_TRANSPORT: Transport = "stdio"
TRANSPORTS: tuple[Transport, ...] = ("stdio", "streamable-http")
MAX_PORT = 65535
LOCALTIME_PATH = "/etc/localtime"
ZONEINFO_MARKER = "zoneinfo/"


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
