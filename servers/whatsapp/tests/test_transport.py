from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

import pytest
from mcp.server.mcpserver import MCPServer
from starlette.testclient import TestClient
from whatsapp_fakes import FakeWaha, make_http, make_settings

from whatsapp_mcp.features.transport.cli import CliOptions, parse_cli
from whatsapp_mcp.features.transport.http import build_http_app
from whatsapp_mcp.features.transport.security import transport_security
from whatsapp_mcp.server import build_server
from whatsapp_mcp.shared.settings import Settings

LOCAL_URL = "http://127.0.0.1:8768"
INITIALIZE = {
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
        "protocolVersion": "2025-06-18",
        "capabilities": {},
        "clientInfo": {"name": "test", "version": "1"},
    },
}
MCP_HEADERS = {"Accept": "application/json, text/event-stream"}


@pytest.fixture
def waha() -> FakeWaha:
    return FakeWaha()


def _server(settings: Settings, waha: FakeWaha) -> MCPServer:
    return build_server(settings, make_http(waha))


def _client(settings: Settings, waha: FakeWaha, base_url: str = LOCAL_URL) -> Iterator[TestClient]:
    with TestClient(build_http_app(_server(settings, waha), settings), base_url=base_url) as client:
        yield client


@pytest.fixture
def open_client(tmp_path: Path, waha: FakeWaha) -> Iterator[TestClient]:
    yield from _client(make_settings(tmp_path), waha)


@pytest.fixture
def token_client(tmp_path: Path, waha: FakeWaha) -> Iterator[TestClient]:
    yield from _client(make_settings(tmp_path, token="s3cret"), waha)


@dataclass(frozen=True)
class Reply:
    status: int
    body: object


def _initialize(client: TestClient, headers: dict[str, str] | None = None) -> Reply:
    response = client.post("/mcp", json=INITIALIZE, headers={**MCP_HEADERS, **(headers or {})})
    content_type = response.headers.get("content-type", "")
    body = response.json() if content_type.startswith("application/json") else response.text
    return Reply(response.status_code, body)


def test_cli_defaults_to_no_overrides() -> None:
    assert parse_cli([]) == CliOptions()


def test_cli_parses_transport_host_and_port() -> None:
    options = parse_cli(["--transport", "stdio", "--host", "0.0.0.0", "--port", "9001"])
    assert options == CliOptions(transport="stdio", host="0.0.0.0", port=9001)


@pytest.mark.parametrize("argv", [["--transport", "sse"], ["--port", "abc"], ["--port", "70000"]])
def test_cli_rejects_bad_values(argv: list[str]) -> None:
    with pytest.raises(SystemExit):
        parse_cli(argv)


def test_health_reports_status_without_token(open_client: TestClient) -> None:
    response = open_client.get("/health")
    assert response.status_code == 200
    assert response.json() == {"status": "WORKING"}


def test_health_never_echoes_upstream_errors(tmp_path: Path, waha: FakeWaha) -> None:
    settings = make_settings(tmp_path, waha_api_key="wrong")
    for client in _client(settings, waha):
        response = client.get("/health")
        assert response.status_code == 503
        assert response.json() == {"status": "UNREACHABLE"}


def test_token_is_required_on_every_route(token_client: TestClient) -> None:
    assert token_client.get("/health").status_code == 401
    missing = _initialize(token_client)
    assert missing == Reply(401, {"error": "unauthorized"})
    assert _initialize(token_client, {"Authorization": "Bearer wrong"}).status == 401
    assert _initialize(token_client, {"Authorization": "Basic s3cret"}).status == 401


def test_valid_token_passes(token_client: TestClient) -> None:
    headers = {"Authorization": "Bearer s3cret"}
    assert token_client.get("/health", headers=headers).json() == {"status": "WORKING"}
    reply = _initialize(token_client, headers)
    assert reply.status == 200
    assert isinstance(reply.body, dict)
    assert reply.body["result"]["serverInfo"]["name"] == "whatsapp"


def test_initialize_works_without_token_when_unset(open_client: TestClient) -> None:
    assert _initialize(open_client).status == 200


def test_foreign_host_is_rejected_by_default(tmp_path: Path, waha: FakeWaha) -> None:
    for client in _client(make_settings(tmp_path), waha, "http://host.docker.internal:8768"):
        assert _initialize(client).status == 421


def test_allowed_hosts_let_foreign_hosts_through(tmp_path: Path, waha: FakeWaha) -> None:
    settings = make_settings(tmp_path, allowed_hosts=("host.docker.internal",))
    for client in _client(settings, waha, "http://host.docker.internal:8768"):
        assert _initialize(client).status == 200
        foreign = _initialize(client, {"Origin": "http://evil.example.com"})
        assert foreign.status == 403


def test_transport_security_keeps_loopback_and_adds_patterns() -> None:
    assert transport_security(()) is None
    settings = transport_security(("host.docker.internal", "proxy.example.com:443", "[fd00::1]"))
    assert settings is not None
    assert settings.enable_dns_rebinding_protection
    assert settings.allowed_hosts == [
        "127.0.0.1:*",
        "localhost:*",
        "[::1]:*",
        "host.docker.internal",
        "host.docker.internal:*",
        "proxy.example.com:443",
        "[fd00::1]",
        "[fd00::1]:*",
    ]
    assert "https://proxy.example.com:443" in settings.allowed_origins
    assert "http://host.docker.internal:*" in settings.allowed_origins
