from __future__ import annotations

from collections.abc import Callable

from office_mcp.shared.events import Shortcut


class RecordingEvents:
    def __init__(self) -> None:
        self.presses: list[tuple[int, Shortcut]] = []
        self.clicks: list[tuple[int, tuple[float, float]]] = []
        self.on_press: Callable[[int, Shortcut], None] | None = None

    def press(self, pid: int, shortcut: Shortcut) -> None:
        self.presses.append((pid, shortcut))
        if self.on_press is not None:
            self.on_press(pid, shortcut)

    def click(self, pid: int, point: tuple[float, float]) -> None:
        self.clicks.append((pid, point))

    def shortcuts(self) -> list[Shortcut]:
        return [shortcut for _, shortcut in self.presses]
