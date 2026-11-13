from __future__ import annotations

import subprocess
import time
from collections.abc import Callable, Iterable
from importlib import import_module
from typing import Any

from office_mcp.shared.errors import OfficeError
from office_mcp.shared.node import Node

ATTRIBUTES = (
    "AXRole",
    "AXTitle",
    "AXValue",
    "AXDescription",
    "AXDOMIdentifier",
    "AXIdentifier",
    "AXDOMClassList",
    "AXChildren",
)
MAX_DEPTH = 90
MAX_NODES = 40000
POLL_SECONDS = 0.5
PERMISSION_HINT = (
    "Accessibility permission is missing for this process. In System Settings > Privacy & "
    "Security > Accessibility, allow the app that starts the server: your terminal (Terminal or "
    "iTerm), Claude.app when Claude Desktop launches it, or /usr/libexec/sshd-keygen-wrapper "
    "when it runs over SSH. Restart the server after granting it."
)


class AxApp:
    def __init__(
        self,
        bundle_id: str,
        app_name: str,
        launch_timeout: float = 15.0,
        sleep: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self.ax: Any = import_module("ApplicationServices")
        self._appkit: Any = import_module("AppKit")
        self._bundle_id = bundle_id
        self._app_name = app_name
        self._launch_timeout = launch_timeout
        self.sleep = sleep
        self.clock = clock

    def trusted(self) -> bool:
        return bool(self.ax.AXIsProcessTrusted())

    def running(self) -> bool:
        return self.pid() is not None

    def require_trust(self) -> None:
        if not self.trusted():
            raise OfficeError(PERMISSION_HINT)

    def pid(self) -> int | None:
        workspace = self._appkit.NSWorkspace.sharedWorkspace()
        for app in workspace.runningApplications():
            if app.bundleIdentifier() == self._bundle_id:
                return int(app.processIdentifier())
        return None

    def application(self, web_content: bool = False) -> Any:
        pid = self.pid()
        if pid is None:
            raise OfficeError(f"{self._app_name} is not running on this Mac. Open it and sign in.")
        application = self.ax.AXUIElementCreateApplication(pid)
        if web_content:
            self.ax.AXUIElementSetAttributeValue(application, "AXManualAccessibility", True)
        return application

    def windows(self, web_content: bool = False) -> list[Any]:
        return list(self.attribute(self.application(web_content), "AXWindows") or [])

    def window(self, matches: Callable[[str], bool], web_content: bool = False) -> Any:
        found = self._matching(matches, web_content)
        if found is None:
            subprocess.run(["open", "-a", self._app_name], check=False, timeout=30)
            deadline = self.clock() + self._launch_timeout
            while found is None and self.clock() < deadline:
                self.sleep(POLL_SECONDS)
                found = self._matching(matches, web_content)
        if found is None:
            raise OfficeError(
                f"{self._app_name} has no matching window, and opening one did not work."
            )
        return found

    def attribute(self, element: Any, name: str) -> Any:
        error, value = self.ax.AXUIElementCopyAttributeValue(element, name, None)
        return value if error == 0 else None

    def set_attribute(self, element: Any, name: str, value: Any) -> None:
        error = self.ax.AXUIElementSetAttributeValue(element, name, value)
        if error != 0:
            raise OfficeError(f"setting {name} failed (error {error})")

    def perform(self, element: Any, action: str, label: str = "") -> None:
        error = self.ax.AXUIElementPerformAction(element, action)
        if error != 0:
            raise OfficeError(f"{action} failed on {label[:80]!r} (error {error})")

    def snapshot(
        self, element: Any, max_depth: int = MAX_DEPTH, skip: Callable[[Node], bool] | None = None
    ) -> Node:
        return Snapshotter(self, max_depth, skip).build(element)

    def values(self, element: Any) -> list[Any]:
        error, values = self.ax.AXUIElementCopyMultipleAttributeValues(element, ATTRIBUTES, 0, None)
        if error == 0 and values is not None and len(values) == len(ATTRIBUTES):
            return list(values)
        return [self.attribute(element, name) for name in ATTRIBUTES]

    def _matching(self, matches: Callable[[str], bool], web_content: bool) -> Any:
        return next(
            (
                w
                for w in self.windows(web_content)
                if matches(as_text(self.attribute(w, "AXTitle")))
            ),
            None,
        )


class Snapshotter:
    def __init__(self, app: AxApp, max_depth: int, skip: Callable[[Node], bool] | None) -> None:
        self._app = app
        self._max_depth = max_depth
        self._skip = skip
        self._budget = MAX_NODES
        self._seen: set[int] = set()

    def build(self, element: Any, depth: int = 0) -> Node:
        self._budget -= 1
        self._seen.add(hash(element))
        role, title, value, description, dom_id, identifier, classes, children = self._app.values(
            element
        )
        node = Node(
            role=as_text(role),
            title=as_text(title),
            value=as_text(value),
            description=as_text(description),
            dom_id=as_text(dom_id) or as_text(identifier),
            classes=tuple(as_text(item) for item in as_sequence(classes)),
            handle=element,
        )
        if depth >= self._max_depth or self._budget <= 0 or (self._skip and self._skip(node)):
            return node
        kids = tuple(
            self.build(child, depth + 1)
            for child in as_sequence(children)
            if hash(child) not in self._seen and self._budget > 0
        )
        return Node(
            role=node.role,
            title=node.title,
            value=node.value,
            description=node.description,
            dom_id=node.dom_id,
            classes=node.classes,
            children=kids,
            handle=element,
        )


def as_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return str(int(value))
    if isinstance(value, int | float):
        return str(value)
    return ""


def as_sequence(value: Any) -> Iterable[Any]:
    if value is None or isinstance(value, str):
        return ()
    try:
        return list(value)
    except TypeError:
        return ()
