from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass, replace
from datetime import tzinfo
from pathlib import Path
from typing import Literal, get_args

from whatsapp_mcp.shared.errors import ConfigurationError
from whatsapp_mcp.shared.timezone import load_zone, system_timezone

Transport = Literal["stdio", "streamable-http"]
TRANSPORTS: tuple[Transport, ...] = get_args(Transport)

DEFAULT_WAHA_URL = "http://127.0.0.1:3500"
DEFAULT_SESSION = "default"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8768
DEFAULT_TRANSPORT: Transport = "streamable-http"
DEFAULT_DRAFTS_PATH = "~/.local/state/whatsapp-mcp/drafts.json"
DEFAULT_DRAFT_TTL_MINUTES = 30
MAX_DRAFT_TTL_MINUTES = 7 * 24 * 60


@dataclass(frozen=True)
class Settings:
    waha_url: str
    waha_api_key: str
    session: str
    host: str
    port: int
    transport: Transport
    token: str | None
    allowed_hosts: tuple[str, ...]
    timezone: tzinfo
    owner_name: str | None
    attachments_dir: Path | None
    drafts_path: Path
    draft_ttl_minutes: int

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> Settings:
        source = os.environ if env is None else env
        reader = _EnvReader(source)
        return cls(
            waha_url=reader.text("WAHA_URL", DEFAULT_WAHA_URL).rstrip("/"),
            waha_api_key=reader.required("WAHA_API_KEY_PLAIN"),
            session=reader.text("WAHA_SESSION", DEFAULT_SESSION),
            host=reader.text("WHATSAPP_MCP_HOST", DEFAULT_HOST),
            port=parse_port(
                reader.text("WHATSAPP_MCP_PORT", str(DEFAULT_PORT)), "WHATSAPP_MCP_PORT"
            ),
            transport=parse_transport(
                reader.text("WHATSAPP_MCP_TRANSPORT", DEFAULT_TRANSPORT), "WHATSAPP_MCP_TRANSPORT"
            ),
            token=reader.optional("WHATSAPP_MCP_TOKEN"),
            allowed_hosts=parse_allowed_hosts(reader.optional("WHATSAPP_MCP_ALLOWED_HOSTS")),
            timezone=_timezone(reader.optional("WHATSAPP_MCP_TIMEZONE"), source),
            owner_name=reader.optional("WHATSAPP_MCP_OWNER_NAME"),
            attachments_dir=_attachments_dir(reader.optional("WHATSAPP_MCP_ATTACHMENTS_DIR")),
            drafts_path=Path(reader.text("WHATSAPP_MCP_DRAFTS", DEFAULT_DRAFTS_PATH)).expanduser(),
            draft_ttl_minutes=parse_bounded_int(
                reader.text("WHATSAPP_MCP_DRAFT_TTL_MINUTES", str(DEFAULT_DRAFT_TTL_MINUTES)),
                "WHATSAPP_MCP_DRAFT_TTL_MINUTES",
                1,
                MAX_DRAFT_TTL_MINUTES,
            ),
        )

    def with_overrides(
        self,
        transport: Transport | None = None,
        host: str | None = None,
        port: int | None = None,
    ) -> Settings:
        return replace(
            self,
            transport=transport or self.transport,
            host=host or self.host,
            port=self.port if port is None else port,
        )


class _EnvReader:
    def __init__(self, env: Mapping[str, str]) -> None:
        self._env = env

    def optional(self, name: str) -> str | None:
        value = self._env.get(name, "").strip()
        return value or None

    def text(self, name: str, default: str) -> str:
        return self.optional(name) or default

    def required(self, name: str) -> str:
        value = self.optional(name)
        if value is None:
            raise ConfigurationError(f"{name} is not set")
        return value


def parse_bounded_int(raw: str, name: str, minimum: int, maximum: int) -> int:
    try:
        value = int(raw.strip())
    except ValueError as error:
        raise ConfigurationError(f"{name} must be a whole number, got {raw!r}") from error
    if not minimum <= value <= maximum:
        raise ConfigurationError(f"{name} must be between {minimum} and {maximum}, got {value}")
    return value


def parse_port(raw: str, name: str) -> int:
    return parse_bounded_int(raw, name, 1, 65_535)


def parse_transport(raw: str, name: str) -> Transport:
    for transport in TRANSPORTS:
        if raw.strip().lower() == transport:
            return transport
    choices = ", ".join(TRANSPORTS)
    raise ConfigurationError(f"{name} must be one of {choices}, got {raw!r}")


def parse_allowed_hosts(raw: str | None) -> tuple[str, ...]:
    if raw is None:
        return ()
    hosts = tuple(entry.strip() for entry in raw.split(",") if entry.strip())
    for host in hosts:
        if "/" in host or any(character.isspace() for character in host):
            raise ConfigurationError(
                f"WHATSAPP_MCP_ALLOWED_HOSTS entries must look like host or host:port, got {host!r}"
            )
    return hosts


def _timezone(raw: str | None, env: Mapping[str, str]) -> tzinfo:
    if raw is None:
        return system_timezone(env)
    return load_zone(raw, "WHATSAPP_MCP_TIMEZONE")


def _attachments_dir(raw: str | None) -> Path | None:
    if raw is None:
        return None
    path = Path(raw).expanduser()
    if not path.is_absolute():
        raise ConfigurationError(
            f"WHATSAPP_MCP_ATTACHMENTS_DIR must be an absolute path, got {raw!r}"
        )
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        raise ConfigurationError(f"WHATSAPP_MCP_ATTACHMENTS_DIR does not exist: {raw!r}") from error
    if not resolved.is_dir():
        raise ConfigurationError(f"WHATSAPP_MCP_ATTACHMENTS_DIR is not a directory: {raw!r}")
    return resolved
