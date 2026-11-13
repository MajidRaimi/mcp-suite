from __future__ import annotations

from collections.abc import Iterator
from dataclasses import replace

import anyio
import pytest
from starlette.testclient import TestClient

from office_mcp.server import apply_arguments, build_parser, build_tools, create_server, http_app
from office_mcp.shared.http import host_patterns, transport_security
from office_mcp.shared.settings import Settings

INITIALIZE = {
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
        "protocolVersion": "2025-06-18",
        "capabilities": {},
        "clientInfo": {"name": "test", "version": "0"},
    },
}
HEADERS = {"Accept": "application/json, text/event-stream"}


def settings(token: str | None = None, allowed_hosts: tuple[str, ...] = ()) -> Settings:
    base = Settings.from_env({"OFFICE_MCP_TIMEZONE": "UTC"})
    return replace(base, token=token, allowed_hosts=allowed_hosts)


def client(current: Settings) -> Iterator[TestClient]:
    app = http_app(create_server(current, build_tools(current)), current)
    with TestClient(app, base_url="http://127.0.0.1:8769") as test_client:
        yield test_client


@pytest.fixture
def open_client() -> Iterator[TestClient]:
    yield from client(settings())


@pytest.fixture
def token_client() -> Iterator[TestClient]:
    yield from client(settings(token="secret", allowed_hosts=("host.docker.internal",)))


def test_every_tool_is_annotated_and_only_reads_that_mark_read_change_state() -> None:
    current = settings()
    server = create_server(current, build_tools(current))

    tools = anyio.run(server.list_tools)

    assert len(tools) == 7
    for tool in tools:
        changes_state = tool.name in {"teams_read", "outlook_read"}
        assert tool.annotations is not None
        assert tool.annotations.read_only_hint is not changes_state
        assert tool.annotations.idempotent_hint is not changes_state
        assert tool.annotations.destructive_hint is False
        assert tool.annotations.open_world_hint is False


def test_http_without_token_answers_initialize(open_client: TestClient) -> None:
    assert open_client.get("/health").status_code == 200
    response = open_client.post("/mcp", json=INITIALIZE, headers=HEADERS)

    assert response.status_code == 200
    assert response.json()["result"]["serverInfo"]["name"] == "office"


def test_token_is_required_on_every_route(token_client: TestClient) -> None:
    assert token_client.get("/health").status_code == 401
    assert token_client.post("/mcp", json=INITIALIZE, headers=HEADERS).status_code == 401
    wrong = {**HEADERS, "Authorization": "Bearer nope"}
    assert token_client.post("/mcp", json=INITIALIZE, headers=wrong).status_code == 401
    right = {**HEADERS, "Authorization": "Bearer secret"}
    assert token_client.get("/health", headers=right).status_code == 200
    assert token_client.post("/mcp", json=INITIALIZE, headers=right).status_code == 200


def test_allowed_hosts_extend_the_loopback_defaults(token_client: TestClient) -> None:
    auth = {**HEADERS, "Authorization": "Bearer secret"}
    docker = {**auth, "Host": "host.docker.internal:8769"}
    evil = {**auth, "Host": "evil.example:8769"}

    assert token_client.post("/mcp", json=INITIALIZE, headers=docker).status_code == 200
    assert token_client.post("/mcp", json=INITIALIZE, headers=evil).status_code == 421


def test_transport_security_defaults_to_sdk_behaviour() -> None:
    assert transport_security(()) is None
    configured = transport_security(("proxy.example:443",))
    assert configured is not None
    assert "proxy.example:443" in configured.allowed_hosts
    assert "https://proxy.example:443" in configured.allowed_origins


def test_padded_token_from_environment_still_authenticates() -> None:
    current = Settings.from_env({"OFFICE_MCP_TIMEZONE": "UTC", "OFFICE_MCP_TOKEN": " secret "})
    app = http_app(create_server(current, build_tools(current)), current)

    with TestClient(app, base_url="http://127.0.0.1:8769") as padded:
        assert padded.get("/health", headers={"Authorization": "Bearer secret"}).status_code == 200


def test_bracketed_ipv6_hosts_get_a_port_wildcard_only_without_a_port() -> None:
    assert host_patterns("[fd00::1]") == ("[fd00::1]", "[fd00::1]:*")
    assert host_patterns("[fd00::1]:8769") == ("[fd00::1]:8769",)
    assert host_patterns("proxy.example") == ("proxy.example", "proxy.example:*")
    assert host_patterns("proxy.example:443") == ("proxy.example:443",)


def test_cli_flags_override_environment() -> None:
    args = build_parser().parse_args(["--http", "--host", "0.0.0.0", "--port", "9100"])

    applied = apply_arguments(settings(), args)

    assert (applied.transport, applied.host, applied.port) == ("streamable-http", "0.0.0.0", 9100)
    assert apply_arguments(settings(), build_parser().parse_args([])).transport == "stdio"
