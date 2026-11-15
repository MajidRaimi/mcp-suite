from __future__ import annotations

import pytest
from fakes.ax_backend import FakeElement, element
from fakes.events import RecordingEvents
from fakes.pasteboard import FakePasteboard
from fakes.trees import TEST_PID, AppRig, app_rig, window

from office_mcp.shared.errors import TextEntryError
from office_mcp.shared.events import Key, Modifier, Shortcut
from office_mcp.shared.pasteboard import CONCEALED_TYPE, PLAIN_TEXT_TYPE, TRANSIENT_TYPE, preserved
from office_mcp.shared.runtime import AppKitPasteboard
from office_mcp.shared.text_input import TextWriter

PLACEHOLDER = "Type a message"
OWNER_CLIPBOARD = [
    {"public.utf8-plain-text": b"owner copy", "public.rtf": b"{\\rtf1 owner}"},
    {"public.png": b"\x89PNG"},
]


def text_box(value: str = PLACEHOLDER + "\n") -> FakeElement:
    box = element("AXTextArea", AXValue=value, AXTitle=PLACEHOLDER)

    def insert(text: object) -> None:
        current = str(box.attributes["AXValue"])
        base = "" if current.rstrip("\n") == PLACEHOLDER else current
        box.attributes["AXValue"] = str(text) + base

    box.on_set["AXSelectedText"] = insert
    return box


def writer_rig(
    box: FakeElement,
) -> tuple[AppRig, TextWriter, RecordingEvents, FakePasteboard]:
    rig = app_rig(window("Chat", box))
    events = RecordingEvents()
    pasteboard = FakePasteboard(OWNER_CLIPBOARD)
    writer = TextWriter(rig.app, events, pasteboard, rig.app.waiter)
    return rig, writer, events, pasteboard


def test_read_treats_the_placeholder_as_empty() -> None:
    box = text_box()
    _, writer, _, _ = writer_rig(box)

    assert writer.read(box, PLACEHOLDER) == ""
    assert writer.read(box) == PLACEHOLDER + "\n"


def test_single_line_text_is_inserted_through_selected_text() -> None:
    box = text_box()
    rig, writer, events, pasteboard = writer_rig(box)

    writer.write(box, "Project Atlas budget is ready", at_start=True, placeholder=PLACEHOLDER)

    assert rig.backend.sets_of("AXFocused") == [True]
    assert rig.backend.sets_of("AXSelectedTextRange") == [(0, 0)]
    assert rig.backend.sets_of("AXSelectedText") == ["Project Atlas budget is ready"]
    assert events.presses == []
    assert pasteboard.writes == []


def test_multi_line_text_is_pasted_and_the_pasteboard_restored() -> None:
    box = text_box()
    rig, writer, events, pasteboard = writer_rig(box)
    text = "Hi Jane,\nthe numbers are in."

    def paste(pid: int, shortcut: Shortcut) -> None:
        clip = pasteboard.text()
        assert clip == text
        assert TRANSIENT_TYPE in pasteboard.items[0]
        box.attributes["AXValue"] = clip

    events.on_press = paste

    writer.write(box, text, placeholder=PLACEHOLDER)

    assert events.presses == [(TEST_PID, Shortcut(Key.V, Modifier.COMMAND))]
    assert rig.backend.sets_of("AXSelectedText") == []
    assert pasteboard.items == OWNER_CLIPBOARD
    assert pasteboard.restores == 1


def test_paste_is_refused_when_the_box_does_not_take_focus() -> None:
    box = text_box()
    box.on_set["AXFocused"] = lambda value: None
    _, writer, events, pasteboard = writer_rig(box)

    with pytest.raises(TextEntryError, match="nothing was pasted"):
        writer.write(box, "Hi Jane,\nthe numbers are in.", placeholder=PLACEHOLDER)

    assert events.presses == []
    assert pasteboard.writes == []


def test_paste_is_refused_while_another_window_stays_in_front() -> None:
    box = text_box()
    rig, writer, events, _ = writer_rig(box)
    other = window("Popped out chat")
    rig.root.add(other)
    rig.root.attributes["AXFocusedWindow"] = other
    chat_window = box.parent
    assert chat_window is not None
    chat_window.actions["AXRaise"] = lambda: None
    chat_window.on_set["AXMain"] = lambda value: None

    with pytest.raises(TextEntryError, match="nothing was pasted"):
        writer.write(box, "Hi Jane,\nthe numbers are in.", placeholder=PLACEHOLDER)

    assert events.presses == []


def test_preserved_restores_after_an_error() -> None:
    pasteboard = FakePasteboard(OWNER_CLIPBOARD)

    try:
        with preserved(pasteboard):
            pasteboard.write_transient_text("temporary")
            raise RuntimeError("boom")
    except RuntimeError:
        pass

    assert pasteboard.items == OWNER_CLIPBOARD


def test_clear_exact_only_clears_its_own_text() -> None:
    box = text_box("Owner draft in progress")
    rig, writer, _, _ = writer_rig(box)

    assert not writer.clear_exact(box, "Project Atlas budget", PLACEHOLDER)
    assert rig.backend.sets_of("AXSelectedText") == []

    box.attributes["AXValue"] = "Project Atlas budget\n"
    box.on_set["AXSelectedText"] = lambda value: box.attributes.update(
        {"AXValue": PLACEHOLDER + "\n"}
    )

    assert writer.clear_exact(box, "Project Atlas budget", PLACEHOLDER)
    assert rig.backend.sets_of("AXSelectedTextRange") == [(0, 21)]


def test_clear_with_keys_selects_its_own_text_and_presses_backspace_once() -> None:
    box = text_box("Project Atlas budget\n")
    rig, writer, events, _ = writer_rig(box)

    def backspace(_pid: int, shortcut: Shortcut) -> None:
        if shortcut == Shortcut(Key.DELETE):
            box.attributes["AXValue"] = PLACEHOLDER + "\n"

    events.on_press = backspace

    assert writer.clear_with_keys(box, "Project Atlas budget", lambda: True, PLACEHOLDER)
    assert rig.backend.sets_of("AXSelectedTextRange") == [(0, 21)]
    assert events.shortcuts() == [Shortcut(Key.DELETE)]
    assert rig.backend.sets_of("AXSelectedText") == []


def test_clear_with_keys_leaves_text_it_did_not_write() -> None:
    box = text_box("Owner draft in progress")
    _, writer, events, _ = writer_rig(box)

    assert not writer.clear_with_keys(box, "Project Atlas budget", lambda: True, PLACEHOLDER)
    assert events.shortcuts() == []


def test_clear_with_keys_selects_nothing_when_the_target_is_already_gone() -> None:
    box = text_box("Project Atlas budget\n")
    rig, writer, events, _ = writer_rig(box)

    assert not writer.clear_with_keys(box, "Project Atlas budget", lambda: False, PLACEHOLDER)
    assert rig.backend.sets_of("AXSelectedTextRange") == []
    assert events.shortcuts() == []


def test_clear_with_keys_presses_nothing_when_the_target_changes_after_select_all() -> None:
    box = text_box("Project Atlas budget\n")
    rig, writer, events, _ = writer_rig(box)
    answers = iter([True, False])

    assert not writer.clear_with_keys(
        box, "Project Atlas budget", lambda: next(answers), PLACEHOLDER
    )
    assert rig.backend.sets_of("AXSelectedTextRange") == [(0, 21)]
    assert events.shortcuts() == []
    assert box.attributes["AXValue"] == "Project Atlas budget\n"


def test_a_paste_that_does_not_land_clears_the_clipboard_and_is_reported() -> None:
    box = text_box()
    _, writer, events, pasteboard = writer_rig(box)

    with pytest.raises(TextEntryError, match="did not land in the text box in time"):
        writer.write(box, "Hi Jane,\nthe numbers are in.", placeholder=PLACEHOLDER)

    assert events.shortcuts() == [Shortcut(Key.V, Modifier.COMMAND)]
    assert pasteboard.items == []
    assert pasteboard.restores == 1
    assert box.attributes["AXValue"] == PLACEHOLDER + "\n"


def test_set_value_writes_ax_value() -> None:
    box = element("AXTextField", AXValue="")
    rig, writer, _, _ = writer_rig(box)

    writer.set_value(box, "jane@example.com")

    assert box.attributes["AXValue"] == "jane@example.com"
    assert rig.backend.sets_of("AXValue") == ["jane@example.com"]


class FakeItem:
    def __init__(self, data: dict[str, bytes] | None = None) -> None:
        self.data: dict[str, bytes] = dict(data or {})

    @classmethod
    def alloc(cls) -> FakeItem:
        return cls()

    def init(self) -> FakeItem:
        return self

    def types(self) -> list[str]:
        return list(self.data)

    def dataForType_(self, kind: str) -> bytes | None:
        return self.data.get(kind)

    def setData_forType_(self, value: bytes, kind: str) -> None:
        self.data[kind] = value

    def setString_forType_(self, value: str, kind: str) -> None:
        self.data[kind] = value.encode()


class FakeBoard:
    def __init__(self, items: list[FakeItem]) -> None:
        self.items = items

    def pasteboardItems(self) -> list[FakeItem]:
        return self.items

    def clearContents(self) -> None:
        self.items = []

    def writeObjects_(self, items: list[FakeItem]) -> None:
        self.items = list(items)


class FakeData:
    @staticmethod
    def dataWithBytes_length_(value: bytes, length: int) -> bytes:
        return value[:length]


class FakeAppKit:
    def __init__(self, board: FakeBoard) -> None:
        self.board = board
        self.NSPasteboardItem = FakeItem
        self.NSData = FakeData
        self.NSPasteboard = self

    def generalPasteboard(self) -> FakeBoard:
        return self.board


def test_appkit_pasteboard_round_trips_every_item_and_type() -> None:
    board = FakeBoard([FakeItem(dict(item)) for item in OWNER_CLIPBOARD])
    pasteboard = AppKitPasteboard(FakeAppKit(board))

    saved = pasteboard.snapshot()
    pasteboard.write_transient_text("Hi Jane,\nall set.")
    written = board.items[0].data

    assert written[PLAIN_TEXT_TYPE] == b"Hi Jane,\nall set."
    assert TRANSIENT_TYPE in written and CONCEALED_TYPE in written
    pasteboard.restore(saved)
    assert [item.data for item in board.items] == OWNER_CLIPBOARD
