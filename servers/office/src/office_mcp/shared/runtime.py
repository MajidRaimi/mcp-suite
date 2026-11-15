from __future__ import annotations

import ctypes
import re
import subprocess
import time
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from datetime import date, datetime
from importlib import import_module
from typing import Any

from office_mcp.shared.ax import AxBackend, ProcessDirectory, PyObjCBackend, as_sequence
from office_mcp.shared.errors import PlatformError
from office_mcp.shared.events import (
    OUTLOOK_FORBIDDEN_SHORTCUTS,
    TEAMS_FORBIDDEN_SHORTCUTS,
    GuardedEvents,
    InputEvents,
    Shortcut,
    event_flags,
)
from office_mcp.shared.focus import Workspace
from office_mcp.shared.pasteboard import (
    CONCEALED_TYPE,
    PLAIN_TEXT_TYPE,
    TRANSIENT_TYPE,
    Pasteboard,
    PasteboardSnapshot,
)
from office_mcp.shared.recipients import RecipientGuard
from office_mcp.shared.settings import Settings
from office_mcp.shared.waiting import Waiter

UNSUPPORTED_PLATFORM = (
    "office-mcp reads Outlook and Teams through the macOS accessibility API, so its tools only "
    "work on macOS with the pyobjc frameworks installed."
)
ACTIVATE_IGNORING_OTHER_APPS = 1 << 1
HID_SYSTEM_STATE = 1
ANY_INPUT_EVENT = 0xFFFFFFFF
LEFT_MOUSE_DOWN = 1
LEFT_MOUSE_UP = 2
LEFT_MOUSE_BUTTON = 0
SCREEN_LOCKED_KEY = "CGSSessionScreenIsLocked"
LSAPPINFO = "/usr/bin/lsappinfo"
LSAPPINFO_TIMEOUT = 2
PID_FIELD = re.compile(r'"pid"\s*=\s*(?P<pid>\d+)')
CARBON_PATH = "/System/Library/Frameworks/Carbon.framework/Carbon"
LAUNCH_COMMAND_TIMEOUT = 30


@dataclass(frozen=True)
class Runtime:
    settings: Settings
    waiter: Waiter
    today: Callable[[], date]
    now: Callable[[], datetime]
    backend: AxBackend
    processes: ProcessDirectory
    workspace: Workspace
    outlook_events: InputEvents
    teams_events: InputEvents
    pasteboard: Pasteboard
    recipients: RecipientGuard


class AppKitProcesses:
    def __init__(self, appkit: Any) -> None:
        self._appkit = appkit

    def pid(self, bundle_id: str) -> int | None:
        workspace = self._appkit.NSWorkspace.sharedWorkspace()
        for app in workspace.runningApplications():
            if app.bundleIdentifier() == bundle_id:
                return int(app.processIdentifier())
        return None

    def launch(self, app_name: str) -> None:
        subprocess.run(["open", "-a", app_name], check=False, timeout=LAUNCH_COMMAND_TIMEOUT)


class AppKitWorkspace:
    def __init__(
        self,
        appkit: Any,
        quartz: Any,
        secure_input: Callable[[], bool],
        front_app: Callable[[], int | None],
    ) -> None:
        self._appkit = appkit
        self._quartz = quartz
        self._secure_input = secure_input
        self._front_app = front_app

    def frontmost_pid(self) -> int | None:
        live = self._front_app()
        if live is not None:
            return live
        front = self._appkit.NSWorkspace.sharedWorkspace().frontmostApplication()
        return None if front is None else int(front.processIdentifier())

    def activate(self, pid: int) -> None:
        app = self._running(pid)
        if app is not None:
            app.activateWithOptions_(ACTIVATE_IGNORING_OTHER_APPS)

    def is_running(self, pid: int) -> bool:
        app = self._running(pid)
        return app is not None and not bool(app.isTerminated())

    def app_name(self, pid: int) -> str:
        app = self._running(pid)
        return "" if app is None else str(app.localizedName() or "")

    def idle_seconds(self) -> float:
        return float(
            self._quartz.CGEventSourceSecondsSinceLastEventType(HID_SYSTEM_STATE, ANY_INPUT_EVENT)
        )

    def screen_locked(self) -> bool:
        session: Mapping[str, object] | None = self._quartz.CGSessionCopyCurrentDictionary()
        return bool(session and session.get(SCREEN_LOCKED_KEY))

    def secure_input(self) -> bool:
        return self._secure_input()

    def _running(self, pid: int) -> Any:
        return self._appkit.NSRunningApplication.runningApplicationWithProcessIdentifier_(pid)


class QuartzEvents:
    def __init__(self, quartz: Any) -> None:
        self._quartz = quartz

    def press(self, pid: int, shortcut: Shortcut) -> None:
        flags = event_flags(shortcut.modifiers)
        for down in (True, False):
            event = self._quartz.CGEventCreateKeyboardEvent(None, int(shortcut.key), down)
            self._quartz.CGEventSetFlags(event, flags)
            self._quartz.CGEventPostToPid(pid, event)

    def click(self, pid: int, point: tuple[float, float]) -> None:
        for kind in (LEFT_MOUSE_DOWN, LEFT_MOUSE_UP):
            event = self._quartz.CGEventCreateMouseEvent(None, kind, point, LEFT_MOUSE_BUTTON)
            self._quartz.CGEventSetFlags(event, 0)
            self._quartz.CGEventPostToPid(pid, event)


class AppKitPasteboard:
    def __init__(self, appkit: Any) -> None:
        self._appkit = appkit

    def snapshot(self) -> PasteboardSnapshot:
        items = self._board().pasteboardItems() or []
        return PasteboardSnapshot(tuple(self._item_data(item) for item in items))

    def restore(self, snapshot: PasteboardSnapshot) -> None:
        board = self._board()
        board.clearContents()
        if snapshot.items:
            board.writeObjects_([self._new_item(item) for item in snapshot.items])

    def write_transient_text(self, text: str) -> None:
        item = self._appkit.NSPasteboardItem.alloc().init()
        item.setString_forType_(text, PLAIN_TEXT_TYPE)
        item.setString_forType_("", TRANSIENT_TYPE)
        item.setString_forType_("", CONCEALED_TYPE)
        board = self._board()
        board.clearContents()
        board.writeObjects_([item])

    def _board(self) -> Any:
        return self._appkit.NSPasteboard.generalPasteboard()

    def _item_data(self, item: Any) -> dict[str, bytes]:
        data: dict[str, bytes] = {}
        for kind in as_sequence(item.types()):
            value = item.dataForType_(kind)
            if value is not None:
                data[str(kind)] = bytes(value)
        return data

    def _new_item(self, data: Mapping[str, bytes]) -> Any:
        item = self._appkit.NSPasteboardItem.alloc().init()
        for kind, value in data.items():
            item.setData_forType_(
                self._appkit.NSData.dataWithBytes_length_(value, len(value)), kind
            )
        return item


def lsappinfo(*arguments: str) -> str:
    try:
        done = subprocess.run(
            [LSAPPINFO, *arguments],
            capture_output=True,
            text=True,
            timeout=LSAPPINFO_TIMEOUT,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return ""
    return done.stdout.strip()


def launch_services_front(run: Callable[..., str] = lsappinfo) -> int | None:
    front = run("front")
    if not front:
        return None
    found = PID_FIELD.search(run("info", "-only", "pid", front))
    return int(found["pid"]) if found else None


def carbon_secure_input() -> Callable[[], bool]:
    carbon = ctypes.cdll.LoadLibrary(CARBON_PATH)
    check = carbon.IsSecureEventInputEnabled
    check.restype = ctypes.c_bool
    return lambda: bool(check())


def build_runtime(settings: Settings) -> Runtime:
    try:
        appkit = import_module("AppKit")
        quartz = import_module("Quartz")
        backend = PyObjCBackend()
        secure_input = carbon_secure_input()
    except (ImportError, OSError) as error:
        raise PlatformError(UNSUPPORTED_PLATFORM) from error
    events = QuartzEvents(quartz)
    return Runtime(
        settings=settings,
        waiter=Waiter(time.monotonic, time.sleep, timeout=settings.action_timeout),
        today=date.today,
        now=lambda: datetime.now(settings.zone.info),
        backend=backend,
        processes=AppKitProcesses(appkit),
        workspace=AppKitWorkspace(appkit, quartz, secure_input, launch_services_front),
        outlook_events=GuardedEvents(events, OUTLOOK_FORBIDDEN_SHORTCUTS),
        teams_events=GuardedEvents(events, TEAMS_FORBIDDEN_SHORTCUTS),
        pasteboard=AppKitPasteboard(appkit),
        recipients=RecipientGuard(settings.send_only_to, settings.self_chat_thread),
    )
