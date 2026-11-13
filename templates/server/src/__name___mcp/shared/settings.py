from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from datetime import UTC, datetime, tzinfo
from enum import StrEnum
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from __name___mcp.shared.errors import ConfigurationError

ENV_PREFIX = "__NAME___MCP_"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8770


class Transport(StrEnum):
    STDIO = "stdio"
    STREAMABLE_HTTP = "streamable-http"


DEFAULT_TRANSPORT = Transport.STDIO


@dataclass(frozen=True)
class Settings:
    transport: Transport
    host: str
    port: int
    token: str | None
    allowed_hosts: tuple[str, ...]
    timezone: tzinfo
    owner_name: str | None

    @classmethod
    def from_env(cls, environ: Mapping[str, str]) -> Settings:
        def read(key: str) -> str | None:
            value = environ.get(f"{ENV_PREFIX}{key}", "").strip()
            return value or None

        return cls(
            transport=parse_transport(read("TRANSPORT") or DEFAULT_TRANSPORT),
            host=read("HOST") or DEFAULT_HOST,
            port=parse_port(read("PORT") or str(DEFAULT_PORT)),
            token=read("TOKEN"),
            allowed_hosts=parse_allowed_hosts(read("ALLOWED_HOSTS")),
            timezone=resolve_timezone(read("TIMEZONE")),
            owner_name=read("OWNER_NAME"),
        )


def parse_transport(raw: str) -> Transport:
    try:
        return Transport(raw)
    except ValueError as error:
        choices = ", ".join(transport.value for transport in Transport)
        raise ConfigurationError(
            f"{ENV_PREFIX}TRANSPORT must be one of {choices}, got {raw!r}"
        ) from error


def parse_port(raw: str) -> int:
    try:
        port = int(raw)
    except ValueError as error:
        raise ConfigurationError(f"{ENV_PREFIX}PORT must be an integer, got {raw!r}") from error
    if not 1 <= port <= 65535:
        raise ConfigurationError(f"{ENV_PREFIX}PORT must be between 1 and 65535, got {port}")
    return port


def parse_allowed_hosts(raw: str | None) -> tuple[str, ...]:
    if raw is None:
        return ()
    return tuple(host.strip() for host in raw.split(",") if host.strip())


def resolve_timezone(raw: str | None) -> tzinfo:
    if raw is None:
        return datetime.now().astimezone().tzinfo or UTC
    try:
        return ZoneInfo(raw)
    except (ZoneInfoNotFoundError, ValueError) as error:
        raise ConfigurationError(
            f"{ENV_PREFIX}TIMEZONE must be an IANA zone such as Europe/Berlin, got {raw!r}"
        ) from error
