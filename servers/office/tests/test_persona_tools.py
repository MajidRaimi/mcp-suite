from __future__ import annotations

from pathlib import Path

import pytest
from fakes.clock import FakeClock
from fakes.lock import MemoryUiLock
from fakes.p7_persona import FULL_FRONT_MATTER, persona_store, write_persona
from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.features.persona.guard import PersonaGuard
from office_mcp.features.persona.store import PersonaStore
from office_mcp.features.persona.tools import PersonaTools
from office_mcp.shared.errors import PersonaFormatError
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import Effect, Stability, ToolGroup
from office_mcp.shared.runner import Runner

OWNER = Owner(name="Alex Example")


def tools_for(store: PersonaStore, lock: MemoryUiLock | None = None) -> PersonaTools:
    runner = Runner(lock or MemoryUiLock(), FakeClock(), lock_timeout=1, stale_after=1)
    return PersonaTools(runner, store, PersonaGuard(store), OWNER)


def loaded(tmp_path: Path) -> PersonaStore:
    return persona_store(write_persona(tmp_path / "persona.md", FULL_FRONT_MATTER))


def test_specs_are_read_tools(tmp_path: Path) -> None:
    specs = tools_for(loaded(tmp_path)).specs()

    assert [spec.name for spec in specs] == ["office_persona", "office_check_text"]
    assert {spec.group for spec in specs} == {ToolGroup.READ}
    assert {spec.effect for spec in specs} == {Effect.READS}
    assert {spec.stability for spec in specs} == {Stability.STABLE}
    assert all("Alex Example" in spec.description for spec in specs)
    assert all(
        "\u2014" not in spec.description and "\u2013" not in spec.description for spec in specs
    )


def test_office_persona_record(tmp_path: Path) -> None:
    record = tools_for(loaded(tmp_path)).office_persona()

    assert record["state"] == "loaded"
    assert record["permissions_ok"] is True
    assert record["enforced"] is True
    assert record["prose"] == "# Voice\nShort and direct."
    assert record["rules"]["channels"]["teams"]["sign_off_rule"] == "forbidden"
    assert set(record["rules_in_words"]) == {"email", "teams", "rsvp"}
    assert {"version", "modified_at", "path", "error"} <= set(record)


def test_office_persona_reports_an_invalid_file_without_failing(tmp_path: Path) -> None:
    store = persona_store(write_persona(tmp_path / "persona.md", "+++\nversion = [\n+++\n"))

    record = tools_for(store).office_persona()

    assert record["state"] == "invalid"
    assert "not valid TOML" in record["error"]
    assert record["rules"] == {}
    assert record["enforced"] is True


def test_office_persona_none_state_is_not_enforced(tmp_path: Path) -> None:
    record = tools_for(persona_store(tmp_path / "persona.md", explicit=False)).office_persona()

    assert record["state"] == "none"
    assert record["enforced"] is False


def test_check_text_returns_the_review_record(tmp_path: Path) -> None:
    record = tools_for(loaded(tmp_path)).office_check_text(
        "The plan \u2014 then.", "email", subject="Project Atlas budget", account="alex@example.com"
    )

    assert record["ok"] is False
    assert record["final_text"] == "The plan \u2014 then.\n\nBest regards,"
    assert record["final_subject"] == "Project Atlas budget"
    assert [item["rule"] for item in record["violations"]] == ["characters"]
    assert "[U+2014]" in record["violations"][0]["message"]
    assert [item["rule"] for item in record["fixes"]] == ["signoff.appended"]
    assert record["persona_version"] == loaded(tmp_path).version


def test_check_text_uses_context_for_language(tmp_path: Path) -> None:
    context = "هل انتهيت من ميزانية المشروع"

    record = tools_for(loaded(tmp_path)).office_check_text("Yes, sent.", "teams", context=context)

    assert [item["rule"] for item in record["violations"]] == ["language.mismatch"]


def test_check_text_reports_a_missing_persona_as_a_violation(tmp_path: Path) -> None:
    store = persona_store(tmp_path / "persona.md", explicit=True)

    record = tools_for(store).office_check_text("Hello", "teams")

    assert record["ok"] is False
    assert record["violations"][0]["rule"] == "persona.unavailable"
    assert "does not exist" in record["violations"][0]["message"]


def test_unlocked_work_still_turns_office_errors_into_tool_errors() -> None:
    runner = Runner(MemoryUiLock(), FakeClock(), lock_timeout=1, stale_after=1)

    def broken() -> None:
        raise PersonaFormatError("the persona file has no front matter")

    with pytest.raises(ToolError, match="no front matter"):
        runner.unlocked(broken)


def test_persona_tools_run_while_another_tool_holds_the_ui_lock(tmp_path: Path) -> None:
    lock = MemoryUiLock()
    lock.available = False
    tools = tools_for(loaded(tmp_path), lock)

    assert tools.office_persona()["state"] == "loaded"
    assert "ok" in tools.office_check_text("Thanks for the update", "teams")
    assert lock.timeouts == []
