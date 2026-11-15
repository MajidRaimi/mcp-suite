from __future__ import annotations

from collections.abc import Callable
from contextlib import suppress

from office_mcp.shared.ax import MENU_OPEN_CODES, AxApp, as_bool, as_int
from office_mcp.shared.drafts import written_exactly
from office_mcp.shared.errors import OfficeError, TextEntryError
from office_mcp.shared.events import InputEvents, Key, Modifier, Shortcut
from office_mcp.shared.pasteboard import CLEARED, Pasteboard
from office_mcp.shared.waiting import Waiter

PASTE = Shortcut(Key.V, Modifier.COMMAND)
BACKSPACE = Shortcut(Key.DELETE)
PASTE_WAIT = 3.0
CLEAR_WAIT = 3.0
LINE_BREAKS = ("\n", "\r")
FOCUS_WAIT = 2.0
WINDOW_CLIMB = 60
NOT_FOCUSED = (
    "the text box did not take keyboard focus in its own window, so nothing was pasted "
    "(another window may be in front)"
)
PASTE_LATE = (
    "the paste did not land in the text box in time, so the clipboard was cleared instead of "
    "restored; check the box before trying again"
)


class TextWriter:
    def __init__(
        self, app: AxApp, events: InputEvents, pasteboard: Pasteboard, waiter: Waiter
    ) -> None:
        self._app = app
        self._events = events
        self._pasteboard = pasteboard
        self._waiter = waiter

    def read(self, element: object, placeholder: str = "") -> str:
        value = self._app.text(element, "AXValue")
        if placeholder and value.rstrip("\n") == placeholder:
            return ""
        return value

    def set_value(self, element: object, text: str) -> None:
        self._app.set_attribute(element, "AXValue", text)

    def write(
        self, element: object, text: str, at_start: bool = False, placeholder: str = ""
    ) -> None:
        self._app.set_attribute(element, "AXFocused", True)
        if at_start:
            self._app.set_attribute(element, "AXSelectedTextRange", self._app.text_range(0, 0))
        if any(mark in text for mark in LINE_BREAKS):
            self.paste(element, text, placeholder)
        else:
            self._app.set_attribute(element, "AXSelectedText", text)

    def clear_exact(self, element: object, inserted: str, placeholder: str = "") -> bool:
        if not written_exactly(inserted, self.read(element, placeholder)):
            return False
        self._app.set_attribute(element, "AXFocused", True)
        self._app.set_attribute(
            element, "AXSelectedTextRange", self._app.text_range(0, self._length(element))
        )
        self._app.set_attribute(element, "AXSelectedText", "")
        return self.read(element, placeholder).strip() == ""

    def clear_with_keys(
        self,
        element: object,
        inserted: str,
        still_valid: Callable[[], bool],
        placeholder: str = "",
    ) -> bool:
        if not written_exactly(inserted, self.read(element, placeholder)):
            return False
        self.focus(element)
        if not still_valid():
            return False
        self._app.set_attribute(
            element, "AXSelectedTextRange", self._app.text_range(0, self._length(element))
        )
        if not still_valid():
            return False
        self._events.press(self._app.require_pid(), BACKSPACE)
        cleared = self._waiter.until(
            lambda: self.read(element, placeholder).strip() == "", CLEAR_WAIT
        )
        return bool(cleared)

    def paste(self, element: object, text: str, placeholder: str = "") -> None:
        self.focus(element)
        before = self.read(element, placeholder)
        saved = self._pasteboard.snapshot()
        restored = saved
        try:
            self._pasteboard.write_transient_text(text)
            self._events.press(self._app.require_pid(), PASTE)
            if not self._landed(element, placeholder, before):
                restored = CLEARED
                raise TextEntryError(PASTE_LATE)
        finally:
            self._pasteboard.restore(restored)

    def _landed(self, element: object, placeholder: str, before: str) -> bool:
        def changed() -> bool:
            return self.read(element, placeholder) != before

        return self._waiter.until(changed, PASTE_WAIT) is not None or changed()

    def press_in(self, element: object, shortcut: Shortcut) -> None:
        self.focus(element)
        self._events.press(self._app.require_pid(), shortcut)

    def focus(self, element: object) -> None:
        window = self._window_of(element)
        if window is None:
            raise TextEntryError(NOT_FOCUSED)
        with suppress(OfficeError):
            self._app.set_attribute(window, "AXMain", True)
        self._app.perform(window, "AXRaise", "text box window", tolerate=MENU_OPEN_CODES)
        self._app.set_attribute(element, "AXFocused", True)
        if self._waiter.until(lambda: self._has_focus(element, window), FOCUS_WAIT) is None:
            raise TextEntryError(NOT_FOCUSED)

    def _has_focus(self, element: object, window: object) -> bool:
        application = self._app.application()
        focused_window = self._app.attribute(application, "AXFocusedWindow")
        if focused_window is None or not self._app.same(focused_window, window):
            return False
        focused = self._app.attribute(application, "AXFocusedUIElement")
        if focused is not None and self._app.same(focused, element):
            return True
        return as_bool(self._app.attribute(element, "AXFocused")) is True

    def _window_of(self, element: object) -> object | None:
        current: object | None = element
        for _ in range(WINDOW_CLIMB):
            if current is None:
                return None
            if self._app.text(current, "AXRole") == "AXWindow":
                return current
            current = self._app.attribute(current, "AXParent")
        return None

    def _length(self, element: object) -> int:
        counted = as_int(self._app.attribute(element, "AXNumberOfCharacters"))
        return counted if counted is not None else len(self._app.text(element, "AXValue"))
