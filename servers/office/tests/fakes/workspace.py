from __future__ import annotations

from collections.abc import Callable


class FakeWorkspace:
    def __init__(
        self,
        frontmost: int | None = 100,
        running: set[int] | None = None,
        names: dict[int, str] | None = None,
    ) -> None:
        self.frontmost = frontmost
        self.running = set(running or set())
        if frontmost is not None:
            self.running.add(frontmost)
        self.names = dict(names or {})
        self.idle: Callable[[], float] = lambda: 60.0
        self.locked = False
        self.secure = False
        self.refuse_activations = 0
        self.activations: list[int] = []
        self.on_activate: Callable[[int], None] | None = None

    def frontmost_pid(self) -> int | None:
        return self.frontmost

    def activate(self, pid: int) -> None:
        self.activations.append(pid)
        if self.refuse_activations > 0:
            self.refuse_activations -= 1
            return
        self.frontmost = pid
        if self.on_activate is not None:
            self.on_activate(pid)

    def is_running(self, pid: int) -> bool:
        return pid in self.running

    def app_name(self, pid: int) -> str:
        return self.names.get(pid, "")

    def idle_seconds(self) -> float:
        return self.idle()

    def screen_locked(self) -> bool:
        return self.locked

    def secure_input(self) -> bool:
        return self.secure
