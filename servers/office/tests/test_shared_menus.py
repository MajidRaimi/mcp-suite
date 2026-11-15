from __future__ import annotations

from collections.abc import Callable

import pytest
from fakes.ax_backend import FakeElement, as_element, element
from fakes.events import RecordingEvents
from fakes.trees import AppRig, app_rig, button, menu_bar, menu_item, popup_menu, window

from office_mcp.shared.errors import (
    ElementNotFoundError,
    ForbiddenActionError,
    MenuItemDisabledError,
    MenuStuckError,
)
from office_mcp.shared.events import Key, Modifier, Shortcut
from office_mcp.shared.menus import (
    MenuBar,
    MenuEntry,
    PopupMenus,
    find_item,
    is_forbidden,
)
from office_mcp.shared.node import Node

NEW_MAIL = ("File", "New", "Mail")


def outlook_bar() -> FakeElement:
    return menu_bar(
        (
            "File",
            [
                menu_item(
                    "New",
                    menu_item("Mail", char="N", modifiers=0),
                    menu_item("Mail", char="N", modifiers=2),
                    menu_item("Event"),
                )
            ],
        ),
        (
            "Message",
            [
                menu_item("Archive", virtual_key=14, modifiers=12),
                menu_item("Delete", virtual_key=51, modifiers=0, enabled=False),
                menu_item("Permanently Delete", virtual_key=51, modifiers=1),
                menu_item("Focused", mark="✓"),
            ],
        ),
    )


def bar_rig() -> AppRig:
    return app_rig(window("Main"), menu_bar=outlook_bar())


def test_entries_read_paths_shortcuts_and_marks() -> None:
    rig = bar_rig()

    entries = {entry.path: entry for entry in MenuBar(rig.app).entries("Message")}

    archive = entries[("Message", "Archive")]
    assert archive.shortcut == Shortcut(Key.E, Modifier.CONTROL)
    assert not entries[("Message", "Delete")].enabled
    assert entries[("Message", "Focused")].checked
    assert not archive.checked


def test_duplicate_titles_need_a_shortcut() -> None:
    rig = bar_rig()
    bar = MenuBar(rig.app)

    with pytest.raises(ElementNotFoundError, match="matches 2 items"):
        bar.entry(NEW_MAIL)
    plain = bar.entry(NEW_MAIL, Shortcut(Key.N, Modifier.COMMAND))
    bar.press(NEW_MAIL, Shortcut(Key.N, Modifier.COMMAND))

    assert plain.shortcut == Shortcut(Key.N, Modifier.COMMAND)
    assert rig.backend.actions_of("AXPress") == [plain.handle]


def test_forbidden_paths_are_refused_before_any_read() -> None:
    rig = bar_rig()
    reads_before = rig.backend.copies

    with pytest.raises(ForbiddenActionError):
        MenuBar(rig.app).press(("Message", "Permanently Delete"))

    assert rig.backend.copies == reads_before
    assert rig.backend.performed == []


@pytest.mark.parametrize(
    "path",
    [
        ("Message", "permanently  delete"),
        ("Mailbox", "Empty Folder"),
        ("Edit", "Toggle Unread"),
        ("Calendar", "Join"),
    ],
)
def test_is_forbidden_matches_whole_phrases_case_insensitively(path: tuple[str, ...]) -> None:
    assert is_forbidden(path)


def test_is_forbidden_does_not_match_partial_words() -> None:
    assert not is_forbidden(("Message", "Delete"))
    assert not is_forbidden(("Message", "Mark as Unread"))
    assert not is_forbidden(("Message", "Joined"))


def test_disabled_leaf_raises_with_frontmost_hint() -> None:
    rig = bar_rig()

    with pytest.raises(MenuItemDisabledError, match="frontmost"):
        MenuBar(rig.app).press(("Message", "Delete"))
    assert rig.backend.performed == []


def test_press_rereads_the_title(monkeypatch: pytest.MonkeyPatch) -> None:
    rig = bar_rig()
    bar = MenuBar(rig.app)
    original = bar.entry

    def renamed(path: tuple[str, ...], shortcut: Shortcut | None = None) -> MenuEntry:
        found = original(path, shortcut)
        as_element(found.handle).attributes["AXTitle"] = "Archive All"
        return found

    monkeypatch.setattr(bar, "entry", renamed)

    with pytest.raises(ElementNotFoundError, match="changed"):
        bar.press(("Message", "Archive"))
    assert rig.backend.performed == []


def test_missing_items_raise_not_found() -> None:
    rig = bar_rig()

    with pytest.raises(ElementNotFoundError):
        MenuBar(rig.app).entry(("Message", "Flag"))
    with pytest.raises(ElementNotFoundError):
        MenuBar(rig.app).entry(("View", "Sidebar"))


def nest(depth: int, leaf: FakeElement) -> FakeElement:
    node = leaf
    for _ in range(depth):
        node = element("AXGroup", node)
    return node


def adds(parent: FakeElement, child: FakeElement, code: int = 0) -> Callable[[], int]:
    def act() -> int:
        parent.add(child)
        return code

    return act


def removes(target: FakeElement) -> Callable[[], int]:
    def act() -> int:
        target.remove()
        return 0

    return act


def popup_rig() -> tuple[AppRig, FakeElement, FakeElement]:
    target = button("More Options")
    other = window("Other", nest(8, element("AXGroup")))
    rig = app_rig(window("Main", nest(5, target)), other)
    return rig, target, other


def popups(rig: AppRig, events: RecordingEvents | None = None) -> PopupMenus:
    return PopupMenus(rig.app, events or RecordingEvents(), rig.app.waiter)


def test_open_finds_a_menu_under_the_element() -> None:
    rig, target, _ = popup_rig()
    menu = popup_menu("Today", "Tomorrow")
    target.actions["AXShowMenu"] = adds(target, menu, -25205)

    opened = popups(rig).open(target)

    assert opened.handle is menu
    assert find_item(opened, ("Tomorrow",)) is not None


def test_open_finds_a_menu_beside_the_element() -> None:
    rig, target, _ = popup_rig()
    menu = popup_menu("Move to Archive")
    holder = target.parent
    assert holder is not None and holder.parent is not None
    target.actions["AXShowMenu"] = adds(holder.parent, menu)

    assert popups(rig).open(target).handle is menu


def test_open_finds_a_menu_in_another_window() -> None:
    rig, target, other = popup_rig()
    menu = popup_menu("Reply with quote")
    deepest = other
    while deepest.children:
        deepest = deepest.children[0]
    target.actions["AXShowMenu"] = adds(deepest, menu)

    assert popups(rig).open(target).handle is menu


def test_open_finds_the_menu_window() -> None:
    rig, target, other = popup_rig()
    cells = element(
        "AXTable",
        element("AXCell", element("AXStaticText", AXValue="Current Folder")),
        element("AXCell", element("AXStaticText", AXValue="All Mailboxes")),
    )
    scope_window = window("menuWindow", cells, AXSubrole="AXUnknown")
    main = rig.root.children[0]

    def show() -> int:
        rig.root.attributes["AXWindows"] = [main, other, scope_window]
        return 0

    target.actions["AXPress"] = show

    opened = popups(rig).open(target, "AXPress")

    assert opened.handle is scope_window
    item = find_item(opened, ("All Mailboxes",))
    assert item is not None and item.role == "AXCell"


def test_open_never_mistakes_the_menu_bar_for_a_popup_near_a_shallow_element() -> None:
    target = button("Search scope")
    bar = menu_bar(("Apple", [menu_item("About This Mac")]))
    main = window("Main", element("AXGroup", target))
    rig = app_rig(main, menu_bar=bar)
    cells = element("AXTable", element("AXCell", element("AXStaticText", AXValue="Mailbox")))
    scope_window = window("menuWindow", cells, AXSubrole="AXUnknown")

    def show() -> int:
        rig.root.attributes["AXWindows"] = [main, scope_window]
        return 0

    target.actions["AXPress"] = show

    opened = popups(rig).open(target, "AXPress")

    assert opened.handle is scope_window


def test_choosing_a_cell_in_a_menu_window_selects_its_row_and_presses_return() -> None:
    rig, target, other = popup_rig()
    wanted = element("AXCell", element("AXStaticText", AXValue="All Mailboxes"))
    row = element("AXRow", wanted, AXSelected=False)
    table = element("AXTable", element("AXRow", element("AXCell", AXValue="Current Folder")), row)
    scope_window = window("menuWindow", table, AXSubrole="AXUnknown")
    main = rig.root.children[0]
    events = RecordingEvents()

    def show() -> int:
        rig.root.attributes["AXWindows"] = [main, other, scope_window]
        return 0

    target.actions["AXPress"] = show
    menus = popups(rig, events)

    menus.choose(menus.open(target, "AXPress"), ("All Mailboxes",))

    assert row.attributes["AXSelected"] is True
    assert events.shortcuts() == [Shortcut(Key.RETURN)]
    assert rig.backend.actions_of("AXPress") == [target]


def test_open_raises_when_no_menu_appears() -> None:
    rig, target, _ = popup_rig()

    with pytest.raises(ElementNotFoundError, match="no menu opened"):
        popups(rig).open(target)


def flag_menu() -> FakeElement:
    return element(
        "AXMenu",
        menu_item("Flag", menu_item("Today"), menu_item("Tomorrow", enabled=False)),
        menu_item("Ignore Conversation"),
        menu_item("Mark as Unread", mark="✓"),
    )


def opened_menu(rig: AppRig, menu: FakeElement) -> Node:
    rig.root.children[0].add(menu)
    return rig.app.snapshot(menu)


def test_choose_presses_a_nested_leaf() -> None:
    rig, _, _ = popup_rig()
    menu = opened_menu(rig, flag_menu())

    popups(rig).choose(menu, ("Flag", "Today"))

    pressed = rig.backend.actions_of("AXPress")
    assert [node.attributes["AXTitle"] for node in pressed] == ["Today"]


def test_choose_refuses_forbidden_and_disabled_items() -> None:
    rig, _, _ = popup_rig()
    menu = opened_menu(rig, flag_menu())
    menus = popups(rig)

    with pytest.raises(ForbiddenActionError):
        menus.choose(menu, ("Ignore Conversation",))
    with pytest.raises(MenuItemDisabledError):
        menus.choose(menu, ("Flag", "Tomorrow"))
    with pytest.raises(ElementNotFoundError):
        menus.choose(menu, ("Categorize",))
    assert rig.backend.performed == []


def test_choose_rereads_the_label_before_pressing() -> None:
    rig, _, _ = popup_rig()
    raw = flag_menu()
    menu = opened_menu(rig, raw)
    raw.children[2].attributes["AXTitle"] = "Mark as Read"

    with pytest.raises(ElementNotFoundError, match="changed"):
        popups(rig).choose(menu, ("Mark as Unread",))


def test_checked_reads_mark_and_toggle_value() -> None:
    rig, _, _ = popup_rig()
    raw = flag_menu()
    raw.add(element("AXCheckBox", AXTitle="Like", AXValue="1"))
    menu = opened_menu(rig, raw)
    menus = popups(rig)

    assert menus.checked(menu, ("Mark as Unread",)) is True
    assert menus.checked(menu, ("Flag", "Today")) is False
    assert menus.checked(menu, ("Like",)) is True
    assert menus.checked(menu, ("Missing",)) is None


def test_close_uses_cancel_first() -> None:
    rig, _, _ = popup_rig()
    raw = popup_menu("Today")
    raw.actions["AXCancel"] = removes(raw)
    menu = opened_menu(rig, raw)
    events = RecordingEvents()

    popups(rig, events).close(menu)

    assert events.presses == []


def test_close_escalates_to_escape_then_raises() -> None:
    rig, _, _ = popup_rig()
    raw = popup_menu("Reply with quote")
    menu = opened_menu(rig, raw)
    events = RecordingEvents()
    events.on_press = lambda pid, shortcut: raw.remove()

    popups(rig, events).close(menu)

    assert events.shortcuts() == [Shortcut(Key.ESCAPE)]
    assert rig.backend.sets_of("AXFocused") == [True]

    stuck = popup_menu("Forward")
    stuck_menu = opened_menu(rig, stuck)
    with pytest.raises(MenuStuckError):
        popups(rig, RecordingEvents()).close(stuck_menu)


def test_close_without_a_menu_finds_open_ones() -> None:
    rig, _, other = popup_rig()
    raw = popup_menu("Today")
    raw.actions["AXCancel"] = removes(raw)
    other.add(raw)

    popups(rig).close()

    assert not raw.alive


def test_a_menu_that_shrinks_to_nothing_counts_as_closed() -> None:
    rig, _, _ = popup_rig()
    raw = flag_menu()
    raw.attributes["AXSize"] = (230.0, 400.0)
    menu = opened_menu(rig, raw)
    events = RecordingEvents()

    def shrink() -> int:
        raw.attributes["AXSize"] = (0.0, 0.0)
        return 0

    raw.actions["AXCancel"] = shrink

    popups(rig, events).close(menu)

    assert events.shortcuts() == []
