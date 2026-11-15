from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass

from office_mcp.shared.ax import AxApp, as_bool
from office_mcp.shared.errors import ElementNotFoundError, UiTimeoutError
from office_mcp.shared.tree import find_all

PRUNED_ROLES = frozenset({"AXTable", "AXOutline", "AXWebArea", "AXList", "AXBrowser"})
LANDMARK_ROLES = frozenset(
    {
        "AXButton",
        "AXGroup",
        "AXTable",
        "AXOutline",
        "AXCheckBox",
        "AXPopUpButton",
        "AXMenuButton",
        "AXRadioButton",
        "AXTextField",
        "AXTextArea",
        "AXComboBox",
        "AXToolbar",
        "AXWebArea",
    }
)
LANDMARK_ATTRIBUTES = ("AXDescription", "AXHelp", "AXTitle")


@dataclass(frozen=True)
class WindowProbe:
    handle: object
    title: str
    subrole: str
    main: bool
    focused: bool
    landmarks: frozenset[str]


class WindowLocator:
    def __init__(
        self, app: AxApp, probe_depth: int = 5, prune_roles: frozenset[str] = PRUNED_ROLES
    ) -> None:
        self._app = app
        self._probe_depth = probe_depth
        self._prune_roles = prune_roles

    def probes(self) -> list[WindowProbe]:
        return [self.probe(window) for window in self._app.windows()]

    def probe(self, window: object) -> WindowProbe:
        return WindowProbe(
            handle=window,
            title=self._app.text(window, "AXTitle"),
            subrole=self._app.text(window, "AXSubrole"),
            main=as_bool(self._app.attribute(window, "AXMain")) is True,
            focused=as_bool(self._app.attribute(window, "AXFocused")) is True,
            landmarks=self._landmarks(window),
        )

    def first[K](self, classify: Callable[[WindowProbe], K], kind: K) -> WindowProbe | None:
        return next((probe for probe in self.probes() if classify(probe) == kind), None)

    def require[K](self, classify: Callable[[WindowProbe], K], kind: K, error: str) -> WindowProbe:
        found = self.first(classify, kind)
        if found is None:
            raise ElementNotFoundError(error)
        return found

    def wait_new[K](
        self, before: Sequence[WindowProbe], classify: Callable[[WindowProbe], K], kind: K
    ) -> WindowProbe:
        return self._app.waiter.require(
            lambda: self._new_window(before, classify, kind),
            lambda: UiTimeoutError(f"no new {kind} window of {self._app.app_name} appeared"),
        )

    def _new_window[K](
        self, before: Sequence[WindowProbe], classify: Callable[[WindowProbe], K], kind: K
    ) -> WindowProbe | None:
        for probe in self.probes():
            known = any(self._app.same(probe.handle, old.handle) for old in before)
            if not known and classify(probe) == kind:
                return probe
        return None

    def _landmarks(self, window: object) -> frozenset[str]:
        elements = find_all(
            self._app,
            window,
            lambda element: self._app.text(element, "AXRole") in LANDMARK_ROLES,
            max_depth=self._probe_depth,
            prune=lambda element: self._app.text(element, "AXRole") in self._prune_roles,
        )
        labels = (
            self._app.text(element, name).strip()
            for element in elements
            for name in LANDMARK_ATTRIBUTES
        )
        return frozenset(label for label in labels if label)
