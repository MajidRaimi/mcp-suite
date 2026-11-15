from __future__ import annotations

from dataclasses import dataclass, field

import pytest
from fakes.events import RecordingEvents

from office_mcp.shared.errors import ForbiddenActionError
from office_mcp.shared.events import (
    FLAG_COMMAND,
    FLAG_SHIFT,
    OUTLOOK_FORBIDDEN_SHORTCUTS,
    TEAMS_FORBIDDEN_SHORTCUTS,
    GuardedEvents,
    Key,
    Modifier,
    Shortcut,
    decode_menu_modifiers,
    menu_key,
)
from office_mcp.shared.runtime import QuartzEvents

CMD = Modifier.COMMAND
SHIFT = Modifier.SHIFT
OPTION = Modifier.OPTION
CONTROL = Modifier.CONTROL


@pytest.mark.parametrize(
    "shortcut",
    [
        Shortcut(Key.DELETE, SHIFT | CMD),
        Shortcut(Key.DELETE, CONTROL),
        Shortcut(Key.J, SHIFT | CMD),
        Shortcut(Key.T, OPTION | CMD),
        Shortcut(Key.Q, CMD),
    ],
)
def test_outlook_forbidden_shortcuts_are_refused(shortcut: Shortcut) -> None:
    inner = RecordingEvents()

    with pytest.raises(ForbiddenActionError, match="never sent"):
        GuardedEvents(inner, OUTLOOK_FORBIDDEN_SHORTCUTS).press(1, shortcut)
    assert inner.presses == []


@pytest.mark.parametrize(
    "shortcut",
    [
        Shortcut(Key.S, OPTION | CMD),
        Shortcut(Key.A, SHIFT | CMD),
        Shortcut(Key.V, SHIFT | CMD),
        Shortcut(Key.D, SHIFT | CMD),
        Shortcut(Key.J, OPTION | CMD),
        Shortcut(Key.O, SHIFT | CMD),
        Shortcut(Key.M, SHIFT | CMD),
        Shortcut(Key.Q, CMD),
    ],
)
def test_teams_forbidden_shortcuts_are_refused(shortcut: Shortcut) -> None:
    inner = RecordingEvents()

    with pytest.raises(ForbiddenActionError):
        GuardedEvents(inner, TEAMS_FORBIDDEN_SHORTCUTS).press(1, shortcut)
    assert inner.presses == []


def test_allowed_shortcuts_and_clicks_pass_through() -> None:
    inner = RecordingEvents()
    guarded = GuardedEvents(inner, OUTLOOK_FORBIDDEN_SHORTCUTS)

    guarded.press(7, Shortcut(Key.DELETE, CMD))
    guarded.press(7, Shortcut(Key.ESCAPE))
    guarded.click(7, (10.0, 20.0))

    assert inner.shortcuts() == [Shortcut(Key.DELETE, CMD), Shortcut(Key.ESCAPE)]
    assert inner.clicks == [(7, (10.0, 20.0))]


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (0, CMD),
        (1, SHIFT | CMD),
        (2, OPTION | CMD),
        (3, SHIFT | OPTION | CMD),
        (4, CONTROL | CMD),
        (8, Modifier.NONE),
        (12, CONTROL),
        (13, SHIFT | CONTROL),
        (24, Modifier.NONE),
        (28, CONTROL),
    ],
)
def test_decode_menu_modifiers(value: int, expected: Modifier) -> None:
    assert decode_menu_modifiers(value) == expected


def test_menu_key_prefers_virtual_key() -> None:
    assert menu_key("", 51) is Key.DELETE
    assert menu_key("n", None) is Key.N
    assert menu_key("[", None) is None
    assert menu_key("", 999) is None


@dataclass
class FakeQuartz:
    calls: list[tuple[str, object]] = field(default_factory=list)

    def CGEventCreateKeyboardEvent(self, source: object, key: int, down: bool) -> object:
        event = ("key", key, down)
        self.calls.append(("create", event))
        return event

    def CGEventCreateMouseEvent(
        self, source: object, kind: int, point: tuple[float, float], button: int
    ) -> object:
        event = ("mouse", kind, point)
        self.calls.append(("create", event))
        return event

    def CGEventSetFlags(self, event: object, flags: int) -> None:
        self.calls.append(("flags", flags))

    def CGEventPostToPid(self, pid: int, event: object) -> None:
        self.calls.append(("post", pid))


def test_quartz_press_sets_flags_explicitly_even_when_zero() -> None:
    quartz = FakeQuartz()
    events = QuartzEvents(quartz)

    events.press(9, Shortcut(Key.ESCAPE))
    events.press(9, Shortcut(Key.V, CMD | SHIFT))

    flags = [value for name, value in quartz.calls if name == "flags"]
    assert flags == [0, 0, FLAG_COMMAND | FLAG_SHIFT, FLAG_COMMAND | FLAG_SHIFT]
    assert [value for name, value in quartz.calls if name == "post"] == [9, 9, 9, 9]


def test_quartz_click_posts_down_and_up_without_modifiers() -> None:
    quartz = FakeQuartz()

    QuartzEvents(quartz).click(9, (5.0, 6.0))

    created = [value for name, value in quartz.calls if name == "create"]
    assert created == [("mouse", 1, (5.0, 6.0)), ("mouse", 2, (5.0, 6.0))]
    assert [value for name, value in quartz.calls if name == "flags"] == [0, 0]
