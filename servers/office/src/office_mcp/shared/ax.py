from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
from enum import StrEnum
from importlib import import_module
from typing import Any, Protocol

from office_mcp.shared.errors import (
    AccessibilityError,
    AccessibilityModeError,
    AppNotRunningError,
    AxActionError,
    UiTimeoutError,
)
from office_mcp.shared.node import Node
from office_mcp.shared.waiting import Waiter

MAX_DEPTH = 90
MAX_NODES = 40000
LAUNCH_TIMEOUT = 15.0
AX_SUCCESS = 0
AX_CANNOT_COMPLETE = -25204
AX_ACTION_UNSUPPORTED = -25205
AX_FAILURE = -25200
AX_NOT_IMPLEMENTED = -25208
MENU_OPEN_CODES = frozenset({AX_FAILURE, AX_CANNOT_COMPLETE, AX_ACTION_UNSUPPORTED})
ENHANCED_CODES = frozenset({AX_NOT_IMPLEMENTED})
CUSTOM_ACTION_NAME = "Name:"
MANUAL_ACCESSIBILITY = "AXManualAccessibility"
ENHANCED_INTERFACE = "AXEnhancedUserInterface"
PERMISSION_HINT = (
    "Accessibility permission is missing for this process. In System Settings > Privacy & "
    "Security > Accessibility, allow the app that starts the server: your terminal (Terminal or "
    "iTerm), Claude.app when Claude Desktop launches it, or /usr/libexec/sshd-keygen-wrapper "
    "when it runs over SSH. Restart the server after granting it."
)


class AxBackend(Protocol):
    def is_trusted(self) -> bool: ...

    def application(self, pid: int) -> object: ...

    def copy(self, element: object, name: str) -> tuple[int, object]: ...

    def copy_many(
        self, element: object, names: Sequence[str]
    ) -> tuple[int, Sequence[object] | None]: ...

    def set(self, element: object, name: str, value: object) -> int: ...

    def perform(self, element: object, action: str) -> int: ...

    def action_names(self, element: object) -> tuple[str, ...]: ...

    def same(self, first: object, second: object) -> bool: ...

    def range_value(self, location: int, length: int) -> object: ...


class ProcessDirectory(Protocol):
    def pid(self, bundle_id: str) -> int | None: ...

    def launch(self, app_name: str) -> None: ...


class AccessMode(StrEnum):
    NATIVE = "native"
    WEB = "web"
    WEB_ENHANCED = "web_enhanced"


class PyObjCBackend:
    def __init__(self) -> None:
        self._ax: Any = import_module("ApplicationServices")
        self._core: Any = import_module("CoreFoundation")

    def is_trusted(self) -> bool:
        return bool(self._ax.AXIsProcessTrusted())

    def application(self, pid: int) -> object:
        return self._ax.AXUIElementCreateApplication(pid)

    def copy(self, element: object, name: str) -> tuple[int, object]:
        error, value = self._ax.AXUIElementCopyAttributeValue(element, name, None)
        return int(error), value

    def copy_many(
        self, element: object, names: Sequence[str]
    ) -> tuple[int, Sequence[object] | None]:
        error, values = self._ax.AXUIElementCopyMultipleAttributeValues(
            element, list(names), 0, None
        )
        return int(error), None if values is None else list(values)

    def set(self, element: object, name: str, value: object) -> int:
        return int(self._ax.AXUIElementSetAttributeValue(element, name, value))

    def perform(self, element: object, action: str) -> int:
        return int(self._ax.AXUIElementPerformAction(element, action))

    def action_names(self, element: object) -> tuple[str, ...]:
        error, names = self._ax.AXUIElementCopyActionNames(element, None)
        if error != AX_SUCCESS or names is None:
            return ()
        return tuple(str(name) for name in names)

    def same(self, first: object, second: object) -> bool:
        return bool(self._core.CFEqual(first, second))

    def range_value(self, location: int, length: int) -> object:
        return self._ax.AXValueCreate(self._ax.kAXValueCFRangeType, (location, length))


class AxApp:
    def __init__(
        self,
        backend: AxBackend,
        processes: ProcessDirectory,
        bundle_id: str,
        app_name: str,
        waiter: Waiter,
        mode: AccessMode = AccessMode.NATIVE,
    ) -> None:
        self.backend = backend
        self.processes = processes
        self.bundle_id = bundle_id
        self.app_name = app_name
        self.waiter = waiter
        self.mode = mode
        self._prepared_pid: int | None = None

    def trusted(self) -> bool:
        return self.backend.is_trusted()

    def require_trust(self) -> None:
        if not self.trusted():
            raise AccessibilityError(PERMISSION_HINT)

    def pid(self) -> int | None:
        return self.processes.pid(self.bundle_id)

    def require_pid(self) -> int:
        pid = self.pid()
        if pid is None:
            raise AppNotRunningError(
                f"{self.app_name} is not running on this Mac. Open it and sign in."
            )
        return pid

    def running(self) -> bool:
        return self.pid() is not None

    def application(self) -> object:
        pid = self.require_pid()
        element = self.backend.application(pid)
        if self.mode is not AccessMode.NATIVE and self._prepared_pid != pid:
            self.backend.set(element, MANUAL_ACCESSIBILITY, True)
            self._prepared_pid = pid
        return element

    def windows(self) -> list[object]:
        return list(as_sequence(self.attribute(self.application(), "AXWindows")))

    def attribute(self, element: object, name: str) -> object | None:
        code, value = self.backend.copy(element, name)
        return value if code == AX_SUCCESS else None

    def attributes(self, element: object, names: Sequence[str]) -> list[object | None]:
        code, values = self.backend.copy_many(element, names)
        if code == AX_SUCCESS and values is not None and len(values) == len(names):
            return list(values)
        return [self.attribute(element, name) for name in names]

    def text(self, element: object, name: str) -> str:
        return as_text(self.attribute(element, name))

    def set_attribute(
        self, element: object, name: str, value: object, tolerate: frozenset[int] = frozenset()
    ) -> None:
        code = self.backend.set(element, name, value)
        if code != AX_SUCCESS and code not in tolerate:
            raise AxActionError(f"setting {name} failed (error {code})")

    def perform(
        self, element: object, action: str, label: str = "", tolerate: frozenset[int] = frozenset()
    ) -> int:
        code = self.backend.perform(element, action)
        if code != AX_SUCCESS and code not in tolerate:
            raise AxActionError(f"{action} failed on {label[:80]!r} (error {code})")
        return code

    def actions(self, element: object) -> tuple[str, ...]:
        return self.backend.action_names(element)

    def custom_action(self, element: object, label: str) -> str | None:
        return next(
            (raw for raw in self.actions(element) if custom_action_label(raw) == label), None
        )

    def snapshot(
        self,
        element: object,
        max_depth: int = MAX_DEPTH,
        prune: Callable[[Node], bool] | None = None,
    ) -> Node:
        from office_mcp.shared.snapshot import Snapshotter

        return Snapshotter(self, max_depth, prune).build(element)

    def same(self, first: object, second: object) -> bool:
        return self.backend.same(first, second)

    def text_range(self, location: int, length: int) -> object:
        return self.backend.range_value(location, length)

    def set_enhanced(self, value: bool) -> bool | None:
        element = self.application()
        before = self.attribute(element, ENHANCED_INTERFACE)
        code = self.backend.set(element, ENHANCED_INTERFACE, value)
        if code != AX_SUCCESS and code not in ENHANCED_CODES:
            raise AccessibilityModeError(
                f"{self.app_name} refused {ENHANCED_INTERFACE}={value} (error {code})"
            )
        after = self.attribute(element, ENHANCED_INTERFACE)
        if after is None or bool(after) != value:
            raise AccessibilityModeError(
                f"{self.app_name} did not apply {ENHANCED_INTERFACE}={value}"
            )
        return None if before is None else bool(before)

    def window_or_launch(
        self, matches: Callable[[object], bool], timeout: float = LAUNCH_TIMEOUT
    ) -> object:
        found = self._matching_window(matches)
        if found is not None:
            return found
        self.processes.launch(self.app_name)
        return self.waiter.require(
            lambda: self._matching_window(matches),
            lambda: UiTimeoutError(
                f"{self.app_name} has no matching window, and opening one did not work."
            ),
            timeout,
        )

    def _matching_window(self, matches: Callable[[object], bool]) -> object | None:
        return next((window for window in self.windows() if matches(window)), None)


def custom_action_label(raw: str) -> str:
    first_line = raw.split("\n", 1)[0]
    if first_line.startswith(CUSTOM_ACTION_NAME):
        return first_line[len(CUSTOM_ACTION_NAME) :].strip()
    return raw


def as_text(value: object) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return str(int(value))
    if isinstance(value, int | float):
        return str(value)
    return ""


def as_bool(value: object) -> bool | None:
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value != 0
    return None


def as_int(value: object) -> int | None:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return None


def as_sequence(value: object) -> Iterable[Any]:
    if value is None or isinstance(value, str | bytes):
        return ()
    if not isinstance(value, Iterable):
        return ()
    try:
        return list(value)
    except TypeError:
        return ()
