from __future__ import annotations

from types import SimpleNamespace

from office_mcp.shared.runtime import AppKitWorkspace, launch_services_front

TEAMS_PID = 502
STALE_PID = 999
FRONT = "ASN:0x0-0xb1fb1f:"


def stale_appkit() -> SimpleNamespace:
    front = SimpleNamespace(processIdentifier=lambda: STALE_PID)
    shared = SimpleNamespace(frontmostApplication=lambda: front)
    return SimpleNamespace(NSWorkspace=SimpleNamespace(sharedWorkspace=lambda: shared))


def workspace(front_app: int | None) -> AppKitWorkspace:
    return AppKitWorkspace(stale_appkit(), SimpleNamespace(), lambda: False, lambda: front_app)


def test_frontmost_comes_from_launch_services_not_the_cached_workspace() -> None:
    assert workspace(TEAMS_PID).frontmost_pid() == TEAMS_PID


def test_without_a_launch_services_answer_the_workspace_answer_is_used() -> None:
    assert workspace(None).frontmost_pid() == STALE_PID


def test_launch_services_front_reads_the_pid_of_the_front_application() -> None:
    calls: list[tuple[str, ...]] = []

    def run(*arguments: str) -> str:
        calls.append(arguments)
        return FRONT if arguments == ("front",) else f'"pid"={TEAMS_PID}'

    assert launch_services_front(run) == TEAMS_PID
    assert calls == [("front",), ("info", "-only", "pid", FRONT)]


def test_launch_services_without_output_gives_no_pid() -> None:
    assert launch_services_front(lambda *_arguments: "") is None
    assert launch_services_front(lambda *arguments: "garbage" if arguments else "") is None
