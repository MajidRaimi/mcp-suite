from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime

from office_mcp.features.outlook.ax import OUTLOOK_APP_NAME, OUTLOOK_BUNDLE, AxOutlook
from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.health import require_working
from office_mcp.features.outlook.rows import RefRegistry, RowFinder
from office_mcp.features.outlook.scope import OutlookScope
from office_mcp.shared.ax import AxApp
from office_mcp.shared.events import InputEvents
from office_mcp.shared.focus import FocusPolicy, FrontHook, Workspace
from office_mcp.shared.menus import MenuBar, PopupMenus
from office_mcp.shared.runtime import Runtime
from office_mcp.shared.scope import ScopeOptions
from office_mcp.shared.text_input import TextWriter
from office_mcp.shared.waiting import Waiter
from office_mcp.shared.windows import WindowLocator


@dataclass(frozen=True)
class OutlookDriver:
    app: AxApp
    gateway: OutlookGateway
    windows: WindowLocator
    menus: MenuBar
    popups: PopupMenus
    text: TextWriter
    events: InputEvents
    workspace: Workspace
    waiter: Waiter
    rows: RowFinder
    registry: RefRegistry
    focus: FocusPolicy
    health: Callable[[], None]
    front: FrontHook

    def scope(self, options: ScopeOptions) -> OutlookScope:
        return OutlookScope(self, options)


def health_check(gateway: OutlookGateway, now: Callable[[], datetime]) -> Callable[[], None]:
    def check() -> None:
        require_working(gateway.shallow_button_labels(), now())

    return check


def build_outlook_driver(runtime: Runtime, registry: RefRegistry) -> OutlookDriver:
    waiter = runtime.waiter
    events = runtime.outlook_events
    app = AxApp(runtime.backend, runtime.processes, OUTLOOK_BUNDLE, OUTLOOK_APP_NAME, waiter)
    windows = WindowLocator(app)
    popups = PopupMenus(app, events, waiter)
    gateway = AxOutlook(app, windows, popups, events)
    settings = runtime.settings
    front = FrontHook()
    return OutlookDriver(
        app=app,
        gateway=gateway,
        windows=windows,
        menus=MenuBar(app),
        popups=popups,
        text=TextWriter(app, events, runtime.pasteboard, waiter),
        events=events,
        workspace=runtime.workspace,
        waiter=waiter,
        rows=RowFinder(gateway, waiter, runtime.today, registry, front),
        registry=registry,
        focus=FocusPolicy(settings.restore_frontmost, settings.idle_seconds),
        health=health_check(gateway, runtime.now),
        front=front,
    )
