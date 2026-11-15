from __future__ import annotations

import os
from pathlib import Path

import pytest
from fakes.p7_persona import FULL_FRONT_MATTER, persona_store, write_persona

from office_mcp.features.persona.factory import build_persona
from office_mcp.features.persona.models import PersonaState
from office_mcp.features.persona.store import content_version
from office_mcp.shared.drafts import NO_PERSONA_VERSION
from office_mcp.shared.errors import PersonaFormatError, PersonaMissingError
from office_mcp.shared.settings import Settings


def bump(path: Path, text: str, step: int) -> None:
    path.write_text(text, encoding="utf-8")
    stamp = path.stat().st_mtime_ns + step * 1_000_000_000
    os.utime(path, ns=(stamp, stamp))


def test_default_path_missing_is_state_none(tmp_path: Path) -> None:
    store = persona_store(tmp_path / "persona.md", explicit=False, required=False)

    status = store.status()

    assert status.state is PersonaState.NONE
    assert status.version == NO_PERSONA_VERSION
    assert store.current().prose == ""
    assert store.current().rules.channels == {}


def test_explicit_path_missing_is_state_missing(tmp_path: Path) -> None:
    store = persona_store(tmp_path / "persona.md", explicit=True, required=False)

    assert store.status().state is PersonaState.MISSING
    with pytest.raises(PersonaMissingError, match="OFFICE_MCP_PERSONA_FILE"):
        store.current()


def test_required_and_missing_is_state_missing(tmp_path: Path) -> None:
    store = persona_store(tmp_path / "persona.md", explicit=False, required=True)

    assert store.status().state is PersonaState.MISSING
    with pytest.raises(PersonaMissingError, match="OFFICE_MCP_PERSONA_REQUIRED"):
        store.current()


def test_invalid_file_is_state_invalid_with_the_error(tmp_path: Path) -> None:
    path = write_persona(tmp_path / "persona.md", "+++\nversion = [\n+++\n")
    store = persona_store(path)

    status = store.status()

    assert status.state is PersonaState.INVALID
    assert status.error is not None and "line 2" in status.error
    assert status.version == content_version(path.read_bytes())
    with pytest.raises(PersonaFormatError):
        store.current()


def test_loaded_file_reports_version_and_time(tmp_path: Path) -> None:
    path = write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER)
    store = persona_store(path)

    status = store.status()

    assert status.state is PersonaState.LOADED
    assert status.version == content_version(FULL_FRONT_MATTER.encode("utf-8"))
    assert len(status.version) == 12
    assert status.modified_at is not None
    assert status.error is None
    assert store.current().prose == "# Voice\nShort and direct."


def test_reloads_when_the_file_changes(tmp_path: Path) -> None:
    path = write_persona(tmp_path / "persona.md", "first voice")
    store = persona_store(path)
    first = store.current()

    bump(path, "second voice, longer", step=5)
    second = store.current()

    assert first.prose == "first voice"
    assert second.prose == "second voice, longer"
    assert first.version != second.version


def test_does_not_reparse_when_unchanged(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER))

    assert store.current() is store.current()


def test_recovers_after_a_fix_and_after_deletion(tmp_path: Path) -> None:
    path = write_persona(tmp_path / "persona.md", "+++\nversion = 9\n+++\n")
    store = persona_store(path, explicit=False)
    assert store.status().state is PersonaState.INVALID

    bump(path, "+++\nversion = 1\n+++\nfixed", step=5)
    assert store.status().state is PersonaState.LOADED

    path.unlink()
    assert store.status().state is PersonaState.NONE


@pytest.mark.parametrize(
    ("mode", "expected"), [(0o600, True), (0o400, True), (0o640, False), (0o604, False)]
)
def test_permission_check(tmp_path: Path, mode: int, expected: bool) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", "voice", mode=mode))

    assert store.status().permissions_ok is expected


def test_status_record_shape(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", "voice"))

    record = store.status().to_record()

    assert set(record) == {"state", "path", "version", "modified_at", "permissions_ok", "error"}
    assert record["state"] == "loaded"
    assert record["path"] == str(tmp_path / "persona.md")


def test_build_persona_uses_settings(tmp_path: Path) -> None:
    settings = Settings.from_env(
        {
            "OFFICE_MCP_PERSONA_FILE": str(tmp_path / "persona.md"),
            "OFFICE_MCP_PERSONA_REQUIRED": "true",
            "OFFICE_MCP_TIMEZONE": "UTC",
        }
    )

    store = build_persona(settings)

    assert store.path == tmp_path / "persona.md"
    assert store.required is True
    assert store.status().state is PersonaState.MISSING
