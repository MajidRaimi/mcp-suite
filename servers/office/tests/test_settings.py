from __future__ import annotations

from pathlib import Path

import pytest

from office_mcp.descriptions import Owner, instructions
from office_mcp.shared.errors import SettingsError
from office_mcp.shared.registry import ToolGroup
from office_mcp.shared.settings import Settings, local_zone


def test_defaults_without_environment() -> None:
    settings = Settings.from_env({})

    assert (settings.host, settings.port, settings.transport) == ("127.0.0.1", 8769, "stdio")
    assert settings.token is None and settings.allowed_hosts == () and settings.owner_name is None
    assert settings.zone.name


def test_reads_every_variable() -> None:
    settings = Settings.from_env(
        {
            "OFFICE_MCP_TIMEZONE": "Europe/Berlin",
            "OFFICE_MCP_HOST": "0.0.0.0",
            "OFFICE_MCP_PORT": "9000",
            "OFFICE_MCP_TRANSPORT": "Streamable-HTTP",
            "OFFICE_MCP_TOKEN": "secret",
            "OFFICE_MCP_ALLOWED_HOSTS": "host.docker.internal, proxy.example:443,",
            "OFFICE_MCP_OWNER_NAME": " Alice ",
        }
    )

    assert settings.zone.name == "Europe/Berlin"
    assert (settings.host, settings.port, settings.transport) == (
        "0.0.0.0",
        9000,
        "streamable-http",
    )
    assert settings.token == "secret"
    assert settings.allowed_hosts == ("host.docker.internal", "proxy.example:443")
    assert settings.owner_name == "Alice"


def test_values_are_stripped_and_blank_means_unset() -> None:
    settings = Settings.from_env(
        {
            "OFFICE_MCP_TIMEZONE": " UTC ",
            "OFFICE_MCP_HOST": "  ",
            "OFFICE_MCP_TOKEN": "secret ",
            "OFFICE_MCP_ALLOWED_HOSTS": " ",
        }
    )

    assert settings.zone.name == "UTC"
    assert settings.host == "127.0.0.1"
    assert settings.token == "secret"
    assert settings.allowed_hosts == ()


def test_local_zone_reads_tz_from_env_then_localtime_link(tmp_path: Path) -> None:
    zoneinfo_dir = tmp_path / "zoneinfo" / "America"
    zoneinfo_dir.mkdir(parents=True)
    (zoneinfo_dir / "Chicago").write_bytes(b"")
    link = tmp_path / "localtime"
    link.symlink_to(zoneinfo_dir / "Chicago")

    assert local_zone({"TZ": "Asia/Tokyo"}, str(link)).name == "Asia/Tokyo"
    assert local_zone({"TZ": ":Asia/Tokyo"}, str(tmp_path / "missing")).name == "Asia/Tokyo"
    assert local_zone({}, str(link)).name == "America/Chicago"
    assert local_zone({"TZ": "Not/AZone"}, str(link)).name == "America/Chicago"


@pytest.mark.parametrize(
    ("variable", "value", "message"),
    [
        ("OFFICE_MCP_PORT", "eighty", "whole number"),
        ("OFFICE_MCP_PORT", "70000", "between 1 and 65535"),
        ("OFFICE_MCP_TRANSPORT", "sse", "must be one of"),
        ("OFFICE_MCP_TIMEZONE", "Mars/Olympus", "not a known IANA zone"),
        ("OFFICE_MCP_ALLOWED_HOSTS", "http://proxy.example", "host or host:port"),
        ("OFFICE_MCP_ALLOWED_HOSTS", "a.example, bad host", "host or host:port"),
    ],
)
def test_invalid_values_raise_settings_error(variable: str, value: str, message: str) -> None:
    with pytest.raises(SettingsError, match=message):
        Settings.from_env({variable: value})


def test_descriptions_use_owner_and_zone() -> None:
    reads = frozenset({ToolGroup.READ})
    assert "the user's Microsoft Outlook" in instructions(Owner(), "UTC", reads)
    assert "Teams are in Europe/Berlin" in instructions(Owner("Alice"), "Europe/Berlin", reads)
    assert "Alice's Microsoft Outlook" in instructions(Owner("Alice"), "UTC", reads)
    assert "tell Alice exactly that" in instructions(Owner("Alice"), "UTC", reads)
