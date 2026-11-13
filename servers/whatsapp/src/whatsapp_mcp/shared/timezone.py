from __future__ import annotations

from collections.abc import Mapping
from datetime import UTC, tzinfo
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from whatsapp_mcp.shared.errors import ConfigurationError

LOCALTIME_PATH = Path("/etc/localtime")
_ZONEINFO_MARKER = "zoneinfo/"


def load_zone(name: str, source: str) -> ZoneInfo:
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError) as error:
        raise ConfigurationError(
            f"{source} must be an IANA time zone such as 'Europe/London', got {name!r}"
        ) from error


def system_timezone(env: Mapping[str, str], localtime: Path = LOCALTIME_PATH) -> tzinfo:
    from_env = _zone_from_tz_variable(env.get("TZ", ""))
    if from_env is not None:
        return from_env
    from_link = _zone_from_localtime(localtime)
    return from_link if from_link is not None else UTC


def _zone_from_tz_variable(raw: str) -> ZoneInfo | None:
    name = raw.strip().removeprefix(":")
    if not name:
        return None
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError):
        return None


def _zone_from_localtime(localtime: Path) -> ZoneInfo | None:
    try:
        target = str(localtime.resolve(strict=True))
    except OSError:
        return None
    _, marker, name = target.partition(_ZONEINFO_MARKER)
    if not marker or not name:
        return None
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError):
        return None
