from __future__ import annotations

import shutil
from collections.abc import Iterator, Mapping
from dataclasses import replace
from pathlib import Path

import anyio
import pytest
from starlette.testclient import TestClient

from office_mcp import container as container_module
from office_mcp.container import all_specs, build_container, exposed_specs
from office_mcp.features.persona.guard import PersonaGuard
from office_mcp.features.persona.models import PersonaState
from office_mcp.server import apply_arguments, build_parser, create_server, http_app
from office_mcp.server import server_instructions as instructions_of
from office_mcp.shared.drafts import Channel, Draft, TextGuard
from office_mcp.shared.errors import SettingsError
from office_mcp.shared.http import host_patterns, transport_security
from office_mcp.shared.registry import EXPERIMENTAL_PREFIX, Stability, annotations_for
from office_mcp.shared.settings import Settings
from office_mcp.tools import settings_record

EXAMPLE_PERSONA = Path(__file__).resolve().parents[1] / "examples" / "persona.example.md"
DASHES = (chr(0x2013), chr(0x2014))
CHECK_FIRST = "Call office_check_text first when unsure."
STABLE_READS = frozenset(
    {
        "office_status",
        "office_persona",
        "office_check_text",
        "outlook_folders",
        "outlook_conversations",
        "outlook_read",
        "outlook_search",
        "outlook_people",
        "outlook_needs_reply",
        "outlook_follow_ups",
        "calendar_events",
        "calendar_event",
        "teams_chats",
        "teams_read",
        "teams_needs_reply",
        "teams_activity",
        "teams_search",
        "teams_channels",
        "teams_channel_read",
    }
)
EXPERIMENTAL = frozenset(
    {
        "outlook_save_attachment",
        "outlook_undo_send",
        "calendar_respond",
        "teams_mark_unread",
    }
)
TEXT_SENDS = (
    "outlook_send",
    "outlook_reply",
    "outlook_forward",
    "calendar_respond",
    "teams_send",
    "teams_reply",
)
EVERYTHING = {"OFFICE_MCP_TOOLS": "all", "OFFICE_MCP_EXPERIMENTAL": "all"}
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


@pytest.fixture(autouse=True)
def isolated_home(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    monkeypatch.setenv("HOME", str(tmp_path))
    return tmp_path


def settings(environ: Mapping[str, str] | None = None) -> Settings:
    return Settings.from_env({"OFFICE_MCP_TIMEZONE": "UTC", **(environ or {})})


def exposed(environ: Mapping[str, str] | None = None) -> set[str]:
    return {spec.name for spec in exposed_specs(build_container(settings(environ)))}


def persona_file(home: Path) -> Path:
    target = home / "persona.md"
    shutil.copyfile(EXAMPLE_PERSONA, target)
    target.chmod(0o600)
    return target


def listed_tools(current: Settings) -> list[str]:
    server = create_server(build_container(current))
    return [tool.name for tool in anyio.run(server.list_tools)]


def client(current: Settings) -> Iterator[TestClient]:
    app = http_app(create_server(build_container(current)), current)
    with TestClient(app, base_url="http://127.0.0.1:8769") as test_client:
        yield test_client


@pytest.fixture
def open_client() -> Iterator[TestClient]:
    yield from client(settings())


@pytest.fixture
def token_client() -> Iterator[TestClient]:
    current = replace(settings(), token="secret", allowed_hosts=("host.docker.internal",))
    yield from client(current)


def test_default_settings_expose_exactly_the_stable_read_tools() -> None:
    assert set(listed_tools(settings())) == STABLE_READS
    assert len(STABLE_READS) == 19


def test_everything_with_a_download_dir_exposes_all_35_tools(tmp_path: Path) -> None:
    names = exposed({**EVERYTHING, "OFFICE_MCP_DOWNLOAD_DIR": str(tmp_path / "downloads")})

    assert len(names) == 35
    assert names >= EXPERIMENTAL


def test_experimental_tools_need_opting_in() -> None:
    names = exposed({"OFFICE_MCP_TOOLS": "all"})

    assert not names & EXPERIMENTAL
    assert "outlook_move" in names
    assert len(names) == 31
    assert "outlook_undo_send" in exposed(
        {"OFFICE_MCP_TOOLS": "all", "OFFICE_MCP_EXPERIMENTAL": "outlook_undo_send"}
    )


def test_read_only_wins_over_tools_and_experimental(tmp_path: Path) -> None:
    environ = {
        **EVERYTHING,
        "OFFICE_MCP_READ_ONLY": "true",
        "OFFICE_MCP_DOWNLOAD_DIR": str(tmp_path / "downloads"),
    }

    assert exposed(environ) == STABLE_READS


def test_disabled_names_are_removed() -> None:
    names = exposed({**EVERYTHING, "OFFICE_MCP_DISABLED_TOOLS": "calendar_respond,outlook_read"})

    assert "calendar_respond" not in names
    assert "outlook_read" not in names
    assert "teams_read" in names


@pytest.mark.parametrize(
    "environ",
    [
        {"OFFICE_MCP_DISABLED_TOOLS": "outlook_nope"},
        {"OFFICE_MCP_EXPERIMENTAL": "teams_teleport"},
    ],
)
def test_unknown_tool_names_stop_startup(environ: dict[str, str]) -> None:
    with pytest.raises(SettingsError, match="unknown tools"):
        create_server(build_container(settings(environ)))


def test_attachment_tool_name_is_known_without_a_download_dir() -> None:
    names = exposed(
        {"OFFICE_MCP_TOOLS": "all", "OFFICE_MCP_EXPERIMENTAL": "outlook_save_attachment"}
    )

    assert "outlook_save_attachment" not in names


def test_every_tool_carries_the_annotations_of_its_effect(tmp_path: Path) -> None:
    current = settings({**EVERYTHING, "OFFICE_MCP_DOWNLOAD_DIR": str(tmp_path / "downloads")})
    container = build_container(current)
    effects = {spec.name: spec.effect for spec in all_specs(container)}

    tools = anyio.run(create_server(container).list_tools)

    assert len(tools) == 35
    for tool in tools:
        assert tool.annotations == annotations_for(effects[tool.name])


def test_experimental_descriptions_carry_the_prefix(tmp_path: Path) -> None:
    container = build_container(settings({"OFFICE_MCP_DOWNLOAD_DIR": str(tmp_path / "dl")}))

    specs = all_specs(container)

    assert {spec.name for spec in specs if spec.stability is Stability.EXPERIMENTAL} == EXPERIMENTAL
    for spec in specs:
        prefixed = spec.description.startswith(EXPERIMENTAL_PREFIX)
        assert prefixed is (spec.stability is Stability.EXPERIMENTAL)


def test_text_sending_descriptions_end_with_the_persona_rules(isolated_home: Path) -> None:
    current = settings({"OFFICE_MCP_PERSONA_FILE": str(persona_file(isolated_home))})
    descriptions = {spec.name: spec.description for spec in all_specs(build_container(current))}

    for name in TEXT_SENDS:
        assert descriptions[name].endswith(CHECK_FIRST), name
        assert descriptions[name].count(CHECK_FIRST) == 1, name
        assert "Persona rules for this text" in descriptions[name], name


def test_persona_block_appears_only_when_send_is_enabled(isolated_home: Path) -> None:
    persona = str(persona_file(isolated_home))
    block = "Everything you write through this server is sent as"

    with_send = instructions_of(
        build_container(settings({"OFFICE_MCP_TOOLS": "all", "OFFICE_MCP_PERSONA_FILE": persona}))
    )
    organize = instructions_of(
        build_container(
            settings({"OFFICE_MCP_TOOLS": "organize", "OFFICE_MCP_PERSONA_FILE": persona})
        )
    )
    reads = instructions_of(build_container(settings({"OFFICE_MCP_PERSONA_FILE": persona})))

    assert block in with_send
    assert block not in organize
    assert block not in reads


def test_only_read_only_instructions_promise_that_nothing_is_sent() -> None:
    reads = instructions_of(build_container(settings()))
    organize = instructions_of(build_container(settings({"OFFICE_MCP_TOOLS": "organize"})))
    everything = instructions_of(build_container(settings(EVERYTHING)))

    assert "nothing is sent" in reads.lower()
    assert "nothing is sent" not in organize.lower()
    assert "nothing is sent" not in everything.lower()
    assert "No persona is configured" in everything


def test_missing_required_persona_blocks_sends_in_instructions(isolated_home: Path) -> None:
    current = settings(
        {
            "OFFICE_MCP_TOOLS": "all",
            "OFFICE_MCP_PERSONA_FILE": str(isolated_home / "absent.md"),
            "OFFICE_MCP_PERSONA_REQUIRED": "true",
        }
    )

    container = build_container(current)

    assert "The persona file is missing" in instructions_of(container)


def test_a_persona_file_created_after_startup_is_enforced_on_sends(
    isolated_home: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    guards: list[TextGuard] = []

    def capture(*args: object) -> object:
        guard = args[-1]
        assert isinstance(guard, PersonaGuard)
        guards.append(guard)
        return object()

    monkeypatch.setattr(container_module, "build_runtime", lambda _settings: object())
    monkeypatch.setattr(container_module, "build_teams_driver", lambda _runtime: object())
    monkeypatch.setattr(container_module, "build_teams_writing", capture)
    container = build_container(settings({"OFFICE_MCP_TOOLS": "all"}))
    assert container.persona.status().state is PersonaState.NONE
    target = isolated_home / ".config" / "office-mcp" / "persona.md"
    target.parent.mkdir(parents=True)
    shutil.copyfile(EXAMPLE_PERSONA, target)
    target.chmod(0o600)

    container.teams_writing.get()
    review = guards[0].review(Draft(channel=Channel.TEAMS, body="plan \u2014 then"))

    assert container.persona.status().state is PersonaState.LOADED
    assert review.ok is False


def test_no_dash_in_any_description_or_instruction(tmp_path: Path) -> None:
    current = settings({**EVERYTHING, "OFFICE_MCP_DOWNLOAD_DIR": str(tmp_path / "downloads")})
    container = build_container(current)
    texts = [spec.description for spec in all_specs(container)]
    texts.append(instructions_of(container))

    for text in texts:
        assert not any(dash in text for dash in DASHES)


def test_status_reports_the_configuration() -> None:
    current = settings(
        {
            "OFFICE_MCP_TOOLS": "organize",
            "OFFICE_MCP_EXPERIMENTAL": "outlook_move",
            "OFFICE_MCP_SEND_ONLY_TO": "jane@example.com,teams:48:notes",
        }
    )

    record = settings_record(current)

    assert record["read_only"] is False
    assert record["tool_groups"] == ["organize", "read"]
    assert record["experimental"] == ["outlook_move"]
    assert record["send_only_to"] == ["jane@example.com", "teams:48:notes"]
    assert record["self_chat_thread"] == "48:notes"
    assert settings_record(settings())["send_only_to"] is None
    assert settings_record(settings({"OFFICE_MCP_READ_ONLY": "1"}))["read_only"] is True


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
    current = settings({"OFFICE_MCP_TOKEN": " secret "})
    app = http_app(create_server(build_container(current)), current)

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
