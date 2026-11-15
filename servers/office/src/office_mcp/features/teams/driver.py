from __future__ import annotations

from dataclasses import dataclass
from datetime import tzinfo

from office_mcp.features.teams.ax import TEAMS_APP_NAME, TEAMS_BUNDLE, AxTeams
from office_mcp.features.teams.chats.opener import ChatOpener
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.features.teams.scope import TeamsScope
from office_mcp.shared.ax import ENHANCED_INTERFACE, AccessMode, AxApp, as_bool
from office_mcp.shared.focus import FocusPolicy, Workspace
from office_mcp.shared.runtime import Runtime
from office_mcp.shared.scope import ScopeOptions
from office_mcp.shared.text_input import TextWriter
from office_mcp.shared.waiting import Waiter


@dataclass(frozen=True)
class TeamsDriver:
    app: AxApp
    gateway: TeamsGateway
    chats: ChatService
    opener: ChatOpener
    text: TextWriter
    workspace: Workspace
    waiter: Waiter
    zone: tzinfo
    focus: FocusPolicy

    def scope(self, options: ScopeOptions) -> TeamsScope:
        return TeamsScope(self, options)


def reset_stale_enhanced(app: AxApp) -> bool:
    if not app.trusted() or not app.running():
        return False
    if as_bool(app.attribute(app.application(), ENHANCED_INTERFACE)) is not True:
        return False
    app.set_enhanced(False)
    return True


def build_teams_driver(runtime: Runtime) -> TeamsDriver:
    waiter = runtime.waiter
    events = runtime.teams_events
    app = AxApp(
        runtime.backend, runtime.processes, TEAMS_BUNDLE, TEAMS_APP_NAME, waiter, AccessMode.WEB
    )
    gateway = AxTeams(app, events, waiter)
    chats = ChatService(gateway, runtime.today, waiter)
    settings = runtime.settings
    reset_stale_enhanced(app)
    return TeamsDriver(
        app=app,
        gateway=gateway,
        chats=chats,
        opener=ChatOpener(gateway, chats, waiter),
        text=TextWriter(app, events, runtime.pasteboard, waiter),
        workspace=runtime.workspace,
        waiter=waiter,
        zone=settings.zone.info,
        focus=FocusPolicy(settings.restore_frontmost, settings.idle_seconds),
    )
