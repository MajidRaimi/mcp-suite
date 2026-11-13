from __future__ import annotations

from datetime import UTC
from pathlib import Path
from zoneinfo import ZoneInfo

import pytest

from whatsapp_mcp.shared.errors import ConfigurationError
from whatsapp_mcp.shared.settings import Settings
from whatsapp_mcp.shared.timezone import system_timezone

BASE = {"WAHA_API_KEY_PLAIN": "secret"}


def _settings(**env: str) -> Settings:
    return Settings.from_env({**BASE, **env})


def test_defaults() -> None:
    settings = _settings(TZ="UTC")
    assert settings.waha_url == "http://127.0.0.1:3500"
    assert settings.session == "default"
    assert settings.host == "127.0.0.1"
    assert settings.port == 8768
    assert settings.transport == "streamable-http"
    assert settings.token is None
    assert settings.allowed_hosts == ()
    assert settings.timezone == ZoneInfo("UTC")
    assert settings.owner_name is None
    assert settings.attachments_dir is None
    assert settings.draft_ttl_minutes == 30


def test_api_key_is_required() -> None:
    with pytest.raises(ConfigurationError, match="WAHA_API_KEY_PLAIN"):
        Settings.from_env({})


@pytest.mark.parametrize("raw", ["abc", "0", "70000", "8768.5"])
def test_port_is_validated(raw: str) -> None:
    with pytest.raises(ConfigurationError, match="WHATSAPP_MCP_PORT"):
        _settings(WHATSAPP_MCP_PORT=raw)


@pytest.mark.parametrize("raw", ["soon", "0", "-5", "999999"])
def test_draft_ttl_is_validated(raw: str) -> None:
    with pytest.raises(ConfigurationError, match="WHATSAPP_MCP_DRAFT_TTL_MINUTES"):
        _settings(WHATSAPP_MCP_DRAFT_TTL_MINUTES=raw)


def test_transport_is_validated_and_normalized() -> None:
    assert _settings(WHATSAPP_MCP_TRANSPORT="STDIO").transport == "stdio"
    with pytest.raises(ConfigurationError, match="WHATSAPP_MCP_TRANSPORT"):
        _settings(WHATSAPP_MCP_TRANSPORT="sse")


def test_optional_values_are_read_and_blank_means_unset() -> None:
    settings = _settings(
        WHATSAPP_MCP_TOKEN="  t0ken  ",
        WHATSAPP_MCP_OWNER_NAME="Alex",
        WHATSAPP_MCP_ALLOWED_HOSTS=" host.docker.internal , ,proxy.example.com:443 ",
    )
    assert settings.token == "t0ken"
    assert settings.owner_name == "Alex"
    assert settings.allowed_hosts == ("host.docker.internal", "proxy.example.com:443")
    blank = _settings(WHATSAPP_MCP_TOKEN="   ", WHATSAPP_MCP_OWNER_NAME="")
    assert blank.token is None
    assert blank.owner_name is None


def test_allowed_hosts_reject_urls() -> None:
    with pytest.raises(ConfigurationError, match="host:port"):
        _settings(WHATSAPP_MCP_ALLOWED_HOSTS="http://proxy.example.com")


def test_timezone_setting_wins_and_is_validated() -> None:
    assert _settings(WHATSAPP_MCP_TIMEZONE="Asia/Tokyo", TZ="UTC").timezone == ZoneInfo(
        "Asia/Tokyo"
    )
    with pytest.raises(ConfigurationError, match="WHATSAPP_MCP_TIMEZONE"):
        _settings(WHATSAPP_MCP_TIMEZONE="Mars/Olympus")


def test_system_timezone_reads_tz_then_localtime_then_falls_back_to_utc(tmp_path: Path) -> None:
    assert system_timezone({"TZ": ":Europe/Paris"}, tmp_path / "none") == ZoneInfo("Europe/Paris")
    zoneinfo_dir = tmp_path / "zoneinfo" / "America"
    zoneinfo_dir.mkdir(parents=True)
    (zoneinfo_dir / "Chicago").write_bytes(b"")
    link = tmp_path / "localtime"
    link.symlink_to(zoneinfo_dir / "Chicago")
    assert system_timezone({}, link) == ZoneInfo("America/Chicago")
    assert system_timezone({}, tmp_path / "missing") is UTC
    assert system_timezone({"TZ": "Not/AZone"}, tmp_path / "missing") is UTC


def test_attachments_dir_must_exist_and_is_resolved(tmp_path: Path) -> None:
    real = tmp_path / "outbox"
    real.mkdir()
    alias = tmp_path / "alias"
    alias.symlink_to(real)
    assert _settings(WHATSAPP_MCP_ATTACHMENTS_DIR=str(alias)).attachments_dir == real.resolve()
    with pytest.raises(ConfigurationError, match="does not exist"):
        _settings(WHATSAPP_MCP_ATTACHMENTS_DIR=str(tmp_path / "missing"))
    with pytest.raises(ConfigurationError, match="absolute"):
        _settings(WHATSAPP_MCP_ATTACHMENTS_DIR="relative/dir")
    file = tmp_path / "file.txt"
    file.write_text("x")
    with pytest.raises(ConfigurationError, match="not a directory"):
        _settings(WHATSAPP_MCP_ATTACHMENTS_DIR=str(file))


def test_cli_overrides_replace_only_given_values() -> None:
    settings = _settings(WHATSAPP_MCP_PORT="9000")
    assert settings.with_overrides().port == 9000
    overridden = settings.with_overrides(transport="stdio", host="0.0.0.0", port=9100)
    assert (overridden.transport, overridden.host, overridden.port) == ("stdio", "0.0.0.0", 9100)
