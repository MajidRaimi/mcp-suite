from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import UTC, date, datetime

from fakes.ax_backend import FakeBackend, FakeProcesses
from fakes.clock import FakeClock, fake_waiter
from fakes.events import RecordingEvents
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.pasteboard import FakePasteboard
from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import AppRig, app_rig, outlook_main_window
from fakes.workspace import FakeWorkspace
from office_mcp.features.outlook.driver import OutlookDriver
from office_mcp.features.outlook.rows import RefRegistry, RowFinder
from office_mcp.features.teams.chats.opener import ChatOpener
from office_mcp.features.teams.chats.service import ChatService
from office_mcp.features.teams.driver import TeamsDriver
from office_mcp.shared.ax import AccessMode
from office_mcp.shared.focus import FocusPolicy, FrontHook
from office_mcp.shared.menus import MenuBar, PopupMenus
from office_mcp.shared.recipients import RecipientGuard
from office_mcp.shared.runtime import Runtime
from office_mcp.shared.settings import Settings
from office_mcp.shared.text_input import TextWriter
from office_mcp.shared.windows import WindowLocator

FAKE_TODAY = date(2026, 10, 20)
FAKE_NOW = datetime(2026, 10, 20, 9, 30, tzinfo=UTC)
NO_IDLE_WAIT = FocusPolicy(restore=True, idle_seconds=0)


@dataclass
class OutlookRig:
    driver: OutlookDriver
    gateway: FakeOutlookGateway
    workspace: FakeWorkspace
    events: RecordingEvents
    app: AppRig
    health_calls: list[str] = field(default_factory=list)


@dataclass
class TeamsRig:
    driver: TeamsDriver
    gateway: FakeTeamsGateway
    workspace: FakeWorkspace
    events: RecordingEvents
    app: AppRig


def fake_settings() -> Settings:
    return Settings.from_env({"OFFICE_MCP_TIMEZONE": "UTC"})


def fake_runtime(
    backend: FakeBackend | None = None,
    processes: FakeProcesses | None = None,
    workspace: FakeWorkspace | None = None,
    clock: FakeClock | None = None,
) -> Runtime:
    settings = fake_settings()
    events = RecordingEvents()
    return Runtime(
        settings=settings,
        waiter=fake_waiter(clock),
        today=lambda: FAKE_TODAY,
        now=lambda: FAKE_NOW,
        backend=backend or FakeBackend(),
        processes=processes or FakeProcesses(),
        workspace=workspace or FakeWorkspace(),
        outlook_events=events,
        teams_events=events,
        pasteboard=FakePasteboard(),
        recipients=RecipientGuard(settings.send_only_to, settings.self_chat_thread),
    )


def outlook_rig(
    gateway: FakeOutlookGateway | None = None,
    health: Callable[[], None] | None = None,
    workspace: FakeWorkspace | None = None,
) -> OutlookRig:
    rig = app_rig(outlook_main_window())
    fake = gateway or FakeOutlookGateway()
    events = RecordingEvents()
    desk = workspace or FakeWorkspace()
    desk.running.add(rig.app.require_pid())
    waiter = rig.app.waiter
    registry = RefRegistry()
    front = FrontHook()
    calls: list[str] = []

    def check() -> None:
        calls.append("health")
        if health is not None:
            health()

    driver = OutlookDriver(
        app=rig.app,
        gateway=fake,
        windows=WindowLocator(rig.app),
        menus=MenuBar(rig.app),
        popups=PopupMenus(rig.app, events, waiter),
        text=TextWriter(rig.app, events, FakePasteboard(), waiter),
        events=events,
        workspace=desk,
        waiter=waiter,
        rows=RowFinder(fake, waiter, lambda: FAKE_TODAY, registry, front),
        registry=registry,
        focus=NO_IDLE_WAIT,
        health=check,
        front=front,
    )
    return OutlookRig(driver, fake, desk, events, rig, calls)


def teams_rig(gateway: FakeTeamsGateway, workspace: FakeWorkspace | None = None) -> TeamsRig:
    rig = app_rig(mode=AccessMode.WEB)
    events = RecordingEvents()
    desk = workspace or FakeWorkspace()
    desk.running.add(rig.app.require_pid())
    waiter = rig.app.waiter
    chats = ChatService(gateway, lambda: FAKE_TODAY, waiter)
    driver = TeamsDriver(
        app=rig.app,
        gateway=gateway,
        chats=chats,
        opener=ChatOpener(gateway, chats, waiter),
        text=TextWriter(rig.app, events, FakePasteboard(), waiter),
        workspace=desk,
        waiter=waiter,
        zone=UTC,
        focus=NO_IDLE_WAIT,
    )
    return TeamsRig(driver, gateway, desk, events, rig)
