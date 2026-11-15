from __future__ import annotations

from dataclasses import dataclass
from enum import Flag, IntEnum, auto
from typing import Protocol

from office_mcp.shared.errors import ForbiddenActionError

MENU_SHIFT = 1
MENU_OPTION = 2
MENU_CONTROL = 4
MENU_NO_COMMAND = 8
FLAG_SHIFT = 0x20000
FLAG_CONTROL = 0x40000
FLAG_OPTION = 0x80000
FLAG_COMMAND = 0x100000


class Key(IntEnum):
    A = 0
    S = 1
    D = 2
    F = 3
    G = 5
    C = 8
    V = 9
    Q = 12
    E = 14
    R = 15
    T = 17
    O = 31
    J = 38
    N = 45
    M = 46
    RETURN = 36
    TAB = 48
    SPACE = 49
    DELETE = 51
    ESCAPE = 53
    LEFT = 123
    RIGHT = 124
    DOWN = 125
    UP = 126


class Modifier(Flag):
    NONE = 0
    COMMAND = auto()
    SHIFT = auto()
    OPTION = auto()
    CONTROL = auto()


MODIFIER_NAMES: tuple[tuple[Modifier, str], ...] = (
    (Modifier.CONTROL, "Control"),
    (Modifier.OPTION, "Option"),
    (Modifier.SHIFT, "Shift"),
    (Modifier.COMMAND, "Cmd"),
)
MODIFIER_FLAGS: tuple[tuple[Modifier, int], ...] = (
    (Modifier.COMMAND, FLAG_COMMAND),
    (Modifier.SHIFT, FLAG_SHIFT),
    (Modifier.OPTION, FLAG_OPTION),
    (Modifier.CONTROL, FLAG_CONTROL),
)


@dataclass(frozen=True)
class Shortcut:
    key: Key
    modifiers: Modifier = Modifier.NONE

    def describe(self) -> str:
        names = [name for modifier, name in MODIFIER_NAMES if modifier in self.modifiers]
        return "+".join([*names, self.key.name.capitalize()])


class InputEvents(Protocol):
    def press(self, pid: int, shortcut: Shortcut) -> None: ...

    def click(self, pid: int, point: tuple[float, float]) -> None: ...


class GuardedEvents:
    def __init__(self, inner: InputEvents, forbidden: frozenset[Shortcut]) -> None:
        self._inner = inner
        self._forbidden = forbidden

    def press(self, pid: int, shortcut: Shortcut) -> None:
        if shortcut in self._forbidden:
            raise ForbiddenActionError(
                f"the shortcut {shortcut.describe()} is never sent by office-mcp"
            )
        self._inner.press(pid, shortcut)

    def click(self, pid: int, point: tuple[float, float]) -> None:
        self._inner.click(pid, point)


def decode_menu_modifiers(value: int) -> Modifier:
    modifiers = Modifier.NONE if value & MENU_NO_COMMAND else Modifier.COMMAND
    if value & MENU_SHIFT:
        modifiers |= Modifier.SHIFT
    if value & MENU_OPTION:
        modifiers |= Modifier.OPTION
    if value & MENU_CONTROL:
        modifiers |= Modifier.CONTROL
    return modifiers


def event_flags(modifiers: Modifier) -> int:
    flags = 0
    for modifier, flag in MODIFIER_FLAGS:
        if modifier in modifiers:
            flags |= flag
    return flags


def menu_key(character: str, virtual_key: int | None) -> Key | None:
    if virtual_key is not None:
        return next((key for key in Key if key.value == virtual_key), None)
    letter = character.strip().upper()
    if len(letter) == 1 and letter.isalpha() and letter in Key.__members__:
        return Key[letter]
    return None


def chord(key: Key, *modifiers: Modifier) -> Shortcut:
    combined = Modifier.NONE
    for modifier in modifiers:
        combined |= modifier
    return Shortcut(key, combined)


OUTLOOK_FORBIDDEN_SHORTCUTS: frozenset[Shortcut] = frozenset(
    {
        chord(Key.DELETE, Modifier.SHIFT, Modifier.COMMAND),
        chord(Key.DELETE, Modifier.CONTROL),
        chord(Key.J, Modifier.SHIFT, Modifier.COMMAND),
        chord(Key.T, Modifier.OPTION, Modifier.COMMAND),
        chord(Key.Q, Modifier.COMMAND),
    }
)
TEAMS_FORBIDDEN_SHORTCUTS: frozenset[Shortcut] = frozenset(
    {
        chord(Key.S, Modifier.OPTION, Modifier.COMMAND),
        chord(Key.A, Modifier.SHIFT, Modifier.COMMAND),
        chord(Key.V, Modifier.SHIFT, Modifier.COMMAND),
        chord(Key.D, Modifier.SHIFT, Modifier.COMMAND),
        chord(Key.J, Modifier.OPTION, Modifier.COMMAND),
        chord(Key.O, Modifier.SHIFT, Modifier.COMMAND),
        chord(Key.M, Modifier.SHIFT, Modifier.COMMAND),
        chord(Key.Q, Modifier.COMMAND),
    }
)
