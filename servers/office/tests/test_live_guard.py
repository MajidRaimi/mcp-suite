from __future__ import annotations

import pytest
from live.guard import LiveRefusal, live_enabled, run_subject, self_targets

STATUS: dict[str, object] = {
    "outlook_account": "Alex@Example.com",
    "self_chat_thread": "48:notes",
    "send_only_to": ["alex@example.com", "teams:48:notes"],
}


def test_live_runs_need_the_flag() -> None:
    assert live_enabled({"OFFICE_MCP_LIVE": "1"})
    assert not live_enabled({})
    assert not live_enabled({"OFFICE_MCP_LIVE": "true"})


def test_self_only_allowlist_is_accepted() -> None:
    targets = self_targets(STATUS, {})

    assert (targets.account, targets.self_thread) == ("alex@example.com", "48:notes")


def test_explicit_self_thread_overrides_the_setting() -> None:
    status = {**STATUS, "send_only_to": ["alex@example.com", "teams:19:abc@unq.gbl.spaces"]}

    targets = self_targets(status, {"OFFICE_MCP_LIVE_SELF_THREAD": "19:abc@unq.gbl.spaces"})

    assert targets.self_thread == "19:abc@unq.gbl.spaces"


@pytest.mark.parametrize(
    "change",
    [
        {"send_only_to": None},
        {"send_only_to": ["alex@example.com"]},
        {"send_only_to": ["@example.com", "teams:48:notes"]},
        {"send_only_to": ["alex@example.com", "teams:48:notes", "jane@example.com"]},
        {"outlook_account": None},
    ],
)
def test_anything_but_self_only_is_refused(change: dict[str, object]) -> None:
    with pytest.raises(LiveRefusal):
        self_targets({**STATUS, **change}, {})


def test_subjects_carry_the_run_prefix() -> None:
    assert run_subject("20261115T120000Z", "T1") == "[mcp-suite test] 20261115T120000Z T1"
