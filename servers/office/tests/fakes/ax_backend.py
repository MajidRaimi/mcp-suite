from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field

AX_SUCCESS = 0
AX_INVALID_ELEMENT = -25202
AX_NO_VALUE = -25212


@dataclass(eq=False)
class FakeElement:
    attributes: dict[str, object] = field(default_factory=dict)
    children: list[FakeElement] = field(default_factory=list)
    parent: FakeElement | None = None
    actions: dict[str, Callable[[], int | None]] = field(default_factory=dict)
    action_codes: dict[str, int] = field(default_factory=dict)
    set_codes: dict[str, int] = field(default_factory=dict)
    on_set: dict[str, Callable[[object], None]] = field(default_factory=dict)
    alive: bool = True
    extra_children: list[FakeElement] = field(default_factory=list)

    def __hash__(self) -> int:
        return 7

    def add(self, *kids: FakeElement) -> FakeElement:
        for kid in kids:
            kid.parent = self
            self.children.append(kid)
        return self

    def remove(self) -> None:
        self.alive = False
        if self.parent is not None and self in self.parent.children:
            self.parent.children.remove(self)


def element(role: str, *kids: FakeElement, **attributes: object) -> FakeElement:
    node = FakeElement(attributes={"AXRole": role, **attributes})
    return node.add(*kids)


class FakeBackend:
    def __init__(self, trusted: bool = True) -> None:
        self.trusted = trusted
        self.apps: dict[int, FakeElement] = {}
        self.performed: list[tuple[FakeElement, str]] = []
        self.sets: list[tuple[FakeElement, str, object]] = []
        self.copies = 0

    def register(self, pid: int, app: FakeElement) -> FakeElement:
        self.apps[pid] = app
        return app

    def is_trusted(self) -> bool:
        return self.trusted

    def application(self, pid: int) -> object:
        return self.apps[pid]

    def copy(self, element: object, name: str) -> tuple[int, object]:
        self.copies += 1
        node = as_element(element)
        if not node.alive:
            return AX_INVALID_ELEMENT, None
        if name == "AXChildren":
            return AX_SUCCESS, [*node.children, *node.extra_children]
        if name == "AXParent":
            return (AX_SUCCESS, node.parent) if node.parent is not None else (AX_NO_VALUE, None)
        if name not in node.attributes:
            return AX_NO_VALUE, None
        return AX_SUCCESS, node.attributes[name]

    def copy_many(
        self, element: object, names: Sequence[str]
    ) -> tuple[int, Sequence[object] | None]:
        node = as_element(element)
        if not node.alive:
            return AX_INVALID_ELEMENT, None
        return AX_SUCCESS, [self.copy(node, name)[1] for name in names]

    def set(self, element: object, name: str, value: object) -> int:
        node = as_element(element)
        self.sets.append((node, name, value))
        code = node.set_codes.get(name, AX_SUCCESS)
        hook = node.on_set.get(name)
        if hook is not None:
            hook(value)
        elif code == AX_SUCCESS:
            node.attributes[name] = value
            self._track_focus(node, name, value)
        return code

    def _track_focus(self, node: FakeElement, name: str, value: object) -> None:
        if value is not True:
            return
        if name == "AXFocused":
            root_of(node).attributes["AXFocusedUIElement"] = node
        if name == "AXMain" and node.attributes.get("AXRole") == "AXWindow":
            root_of(node).attributes["AXFocusedWindow"] = node

    def perform(self, element: object, action: str) -> int:
        node = as_element(element)
        self.performed.append((node, action))
        callback = node.actions.get(action)
        result = callback() if callback is not None else None
        if callback is None and action == "AXRaise" and node.attributes.get("AXRole") == "AXWindow":
            root_of(node).attributes["AXFocusedWindow"] = node
        if result is not None:
            return result
        return node.action_codes.get(action, AX_SUCCESS)

    def action_names(self, element: object) -> tuple[str, ...]:
        names = as_element(element).attributes.get("actions", ())
        return tuple(str(name) for name in names) if isinstance(names, tuple) else ()

    def same(self, first: object, second: object) -> bool:
        return first is second

    def range_value(self, location: int, length: int) -> object:
        return (location, length)

    def sets_of(self, name: str) -> list[object]:
        return [value for _, attribute, value in self.sets if attribute == name]

    def actions_of(self, action: str) -> list[FakeElement]:
        return [node for node, name in self.performed if name == action]


class FakeProcesses:
    def __init__(self, pids: dict[str, int] | None = None) -> None:
        self.pids = dict(pids or {})
        self.launched: list[str] = []
        self.on_launch: Callable[[str], None] | None = None

    def pid(self, bundle_id: str) -> int | None:
        return self.pids.get(bundle_id)

    def launch(self, app_name: str) -> None:
        self.launched.append(app_name)
        if self.on_launch is not None:
            self.on_launch(app_name)


def root_of(node: FakeElement) -> FakeElement:
    current = node
    while current.parent is not None:
        current = current.parent
    return current


def as_element(value: object) -> FakeElement:
    if not isinstance(value, FakeElement):
        raise TypeError(f"not a fake element: {value!r}")
    return value
