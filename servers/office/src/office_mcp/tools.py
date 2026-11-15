from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from office_mcp import descriptions as text
from office_mcp.features.outlook.driver import OutlookDriver
from office_mcp.features.teams.driver import TeamsDriver
from office_mcp.shared.ax import ENHANCED_INTERFACE, AxApp, as_bool
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.focus import Workspace
from office_mcp.shared.registry import Effect, ToolGroup, ToolPolicy, ToolSpec
from office_mcp.shared.settings import Settings

if TYPE_CHECKING:
    from office_mcp.container import Container

Record = dict[str, Any]
ALL_EXPERIMENTAL = "all"


def guarded[T](read: Callable[[], T]) -> T | None:
    try:
        return read()
    except OfficeError:
        return None


def policy_record(policy: ToolPolicy) -> Record:
    experimental = [ALL_EXPERIMENTAL] if policy.all_experimental else sorted(policy.experimental)
    return {
        "read_only": policy.groups <= {ToolGroup.READ},
        "tool_groups": sorted(group.value for group in policy.groups),
        "experimental": experimental,
        "disabled_tools": sorted(policy.disabled),
    }


def settings_record(settings: Settings) -> Record:
    allowlist = settings.send_only_to
    return {
        **policy_record(settings.tool_policy),
        "send_only_to": None if allowlist is None else list(allowlist.entries()),
        "self_chat_thread": settings.self_chat_thread,
    }


def session_state(workspace: Workspace) -> str:
    if workspace.screen_locked():
        return "locked"
    if workspace.secure_input():
        return "secure_input"
    return "ok"


def frontmost_app(workspace: Workspace) -> str | None:
    pid = workspace.frontmost_pid()
    return None if pid is None else workspace.app_name(pid)


def enhanced_mode(app: AxApp) -> bool | None:
    if not (app.trusted() and app.running()):
        return None
    return guarded(lambda: as_bool(app.attribute(app.application(), ENHANCED_INTERFACE)))


def outlook_record(outlook: OutlookDriver, health: Callable[[], Record]) -> Record:
    running = outlook.app.running()
    record: Record = {"outlook_running": running}
    if not (running and outlook.app.trusted()):
        return record
    return {
        **record,
        **health(),
        "outlook_account": guarded(outlook.gateway.account),
        "outlook_mode": guarded(lambda: outlook.gateway.mode().value),
    }


def apps_record(outlook: OutlookDriver, teams: TeamsDriver, health: Callable[[], Record]) -> Record:
    workspace = outlook.workspace
    return {
        "accessibility": outlook.app.trusted(),
        **outlook_record(outlook, health),
        "teams_running": teams.app.running(),
        "teams_enhanced_mode": enhanced_mode(teams.app),
        "frontmost_app": guarded(lambda: frontmost_app(workspace)),
        "session": session_state(workspace),
    }


def office_status(container: Container) -> Record:
    return {
        **apps_record(
            container.outlook.get(),
            container.teams.get(),
            lambda: container.mail.get().health_record(),
        ),
        **settings_record(container.settings),
        "persona": container.persona.status().to_record(),
    }


class OfficeTools:
    def __init__(self, container: Container) -> None:
        self._container = container

    def specs(self) -> list[ToolSpec]:
        return [
            ToolSpec(
                name="office_status",
                function=self.office_status,
                description=text.OFFICE_STATUS,
                group=ToolGroup.READ,
                effect=Effect.READS,
            )
        ]

    def office_status(self) -> Record:
        return self._container.runner.read(lambda _deadline: office_status(self._container))
