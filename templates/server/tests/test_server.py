from __future__ import annotations

from datetime import UTC, datetime
from zoneinfo import ZoneInfo

import anyio
import pytest
from starlette.types import ASGIApp, Message, Receive, Scope, Send

from __name___mcp.features.status.service import StatusService
from __name___mcp.server import apply_overrides, build_http_app, build_server, parse_args
from __name___mcp.shared.errors import ConfigurationError
from __name___mcp.shared.http import BearerTokenMiddleware, transport_security
from __name___mcp.shared.settings import DEFAULT_PORT, Settings, Transport

PREFIX = "__NAME___MCP_"


def make_settings(**overrides: str) -> Settings:
    return Settings.from_env({f"{PREFIX}{key}": value for key, value in overrides.items()})


def test_defaults_are_loopback_stdio_without_token() -> None:
    settings = make_settings(TIMEZONE="UTC")
    assert settings.transport is Transport.STDIO
    assert settings.host == "127.0.0.1"
    assert settings.port == DEFAULT_PORT
    assert settings.token is None
    assert settings.allowed_hosts == ()
    assert settings.owner_name is None


@pytest.mark.parametrize("port", ["eighty", "0", "70000"])
def test_invalid_port_raises_configuration_error(port: str) -> None:
    with pytest.raises(ConfigurationError, match="PORT"):
        make_settings(PORT=port)


def test_invalid_transport_and_timezone_raise_configuration_error() -> None:
    with pytest.raises(ConfigurationError, match="TRANSPORT"):
        make_settings(TRANSPORT="carrier-pigeon")
    with pytest.raises(ConfigurationError, match="TIMEZONE"):
        make_settings(TIMEZONE="Mars/Olympus_Mons")


def test_cli_flags_override_environment() -> None:
    settings = make_settings(TRANSPORT="stdio", PORT="9000")
    args = parse_args(["--transport", "streamable-http", "--port", "9100", "--host", "0.0.0.0"])
    merged = apply_overrides(settings, args)
    assert merged.transport is Transport.STREAMABLE_HTTP
    assert merged.port == 9100
    assert merged.host == "0.0.0.0"


def test_allowed_hosts_extend_loopback_defaults() -> None:
    security = transport_security(("host.docker.internal", "mcp.example.test:443"))
    assert security is not None
    assert "127.0.0.1:*" in security.allowed_hosts
    assert "host.docker.internal:*" in security.allowed_hosts
    assert "mcp.example.test:443" in security.allowed_hosts
    assert "http://host.docker.internal:*" in security.allowed_origins
    assert transport_security(()) is None


def test_status_reports_time_in_injected_timezone() -> None:
    service = StatusService(
        "Test Owner",
        ZoneInfo("America/New_York"),
        clock=lambda: datetime(2026, 1, 15, 12, 0, tzinfo=UTC),
    )
    report = service.report()
    assert report.owner == "Test Owner"
    assert report.timezone == "America/New_York"
    assert report.local_time == "2026-01-15T07:00:00-05:00"


def test_every_tool_declares_annotations_and_owner_fallback() -> None:
    server = build_server(make_settings(TIMEZONE="UTC"))
    tools = anyio.run(server.list_tools)
    assert tools
    for tool in tools:
        assert tool.annotations is not None
        assert tool.annotations.read_only_hint is not None
        assert "the user" in (tool.description or "")


async def call_app(app: ASGIApp, headers: list[tuple[bytes, bytes]]) -> int:
    scope: Scope = {
        "type": "http",
        "asgi": {"version": "3.0"},
        "http_version": "1.1",
        "method": "GET",
        "scheme": "http",
        "path": "/health",
        "raw_path": b"/health",
        "query_string": b"",
        "root_path": "",
        "headers": [(b"host", b"127.0.0.1:8770"), *headers],
        "client": ("127.0.0.1", 50000),
        "server": ("127.0.0.1", 8770),
    }
    messages: list[Message] = []

    async def receive() -> Message:
        return {"type": "http.request", "body": b"", "more_body": False}

    async def send(message: Message) -> None:
        messages.append(message)

    await app(scope, receive, send)
    start = next(message for message in messages if message["type"] == "http.response.start")
    status: int = start["status"]
    return status


async def reached_app(scope: Scope, receive: Receive, send: Send) -> None:
    raise AssertionError("request should have been rejected")


@pytest.mark.parametrize(
    ("headers", "expected"),
    [
        ([], 401),
        ([(b"authorization", b"Bearer wrong")], 401),
        ([(b"authorization", b"Bearer s3cret")], 200),
    ],
)
def test_token_guards_every_http_route(headers: list[tuple[bytes, bytes]], expected: int) -> None:
    settings = make_settings(TIMEZONE="UTC", TOKEN="s3cret")
    app = build_http_app(build_server(settings), settings)
    assert anyio.run(call_app, app, headers) == expected


def test_token_middleware_rejects_before_reaching_app() -> None:
    middleware = BearerTokenMiddleware(reached_app, token="s3cret")
    assert anyio.run(call_app, middleware, []) == 401
