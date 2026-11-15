from __future__ import annotations

from dataclasses import dataclass

import pytest
from fakes.ax_backend import FakeElement, element
from fakes.clock import fake_waiter
from fakes.events import RecordingEvents
from fakes.trees import AppRig, app_rig, outlook_main_window, window

from office_mcp.features.outlook.ax import AxOutlook
from office_mcp.features.outlook.rows import SearchScope, SearchState
from office_mcp.features.outlook.suggestions import SUGGESTIONS_WINDOW
from office_mcp.features.outlook.window import FocusedTab, OutlookMode, OutlookModule
from office_mcp.shared.errors import (
    CleanupError,
    ElementNotFoundError,
    ForbiddenActionError,
    UiTimeoutError,
)
from office_mcp.shared.events import Key, Shortcut
from office_mcp.shared.menus import PopupMenus
from office_mcp.shared.node import Node
from office_mcp.shared.windows import WindowLocator

ARCHIVE_ACTION = "Name:Archive\nTarget:0x0\nSelector:(null)"
SETTINGS_ACTION = "Name:Settings...\nTarget:0x0\nSelector:(null)"
INBOX_TITLE = "Inbox • jane@example.com"
SEARCH_TITLE = 'Searching "jane@example.com"'


@dataclass
class OutlookScreen:
    rig: AppRig
    outlook: AxOutlook
    main: FakeElement
    title_bar: FakeElement
    toggle: FakeElement
    rows: list[FakeElement]
    field: FakeElement
    events: RecordingEvents


def cell_row(description: str, **attributes: object) -> FakeElement:
    return element("AXRow", element("AXCell", AXDescription=description), **attributes)


def screen() -> OutlookScreen:
    rows = [
        cell_row("  Alice Example, Project Atlas budget,  9:15 AM,  Numbers", AXSelected=True),
        cell_row("  Bob Sample, Atlas review,  Yesterday,  See you", actions=(ARCHIVE_ACTION,)),
    ]
    folders = [
        cell_row("jane@example.com", AXDisclosureLevel=0),
        cell_row("Inbox; 2 unread messages", AXDisclosureLevel=1),
    ]
    toggle = element("AXCheckBox", AXTitle="Focused Inbox Selected;  Other Inbox Unselected")
    toggle.attributes["AXValue"] = 0
    scope = element("AXPopUpButton", AXDescription="Search Scope", AXValue="Current Mailbox")
    perform = element("AXButton", AXDescription="Perform Search")
    field = element(
        "AXTextField", scope, perform, AXDescription="Search", AXSubrole="AXSearchField"
    )
    title_bar = element("AXGroup", field, AXTitle="Title bar")
    message_list = element(
        "AXGroup",
        element("AXGroup", toggle, AXDescription="Message List Header"),
        element("AXTable", *rows, AXDescription="Message List", AXRows=rows),
        AXDescription="Message List",
    )
    outline = element("AXOutline", *folders, AXRows=folders)
    calendar_button = element("AXButton", AXHelp="Calendar")
    main = outlook_main_window(
        title_bar,
        element("AXSplitGroup", element("AXScrollArea", outline), message_list),
        title=INBOX_TITLE,
    )
    main.children[0].children[1] = calendar_button
    calendar_button.parent = main.children[0]
    compose = window(
        "Untitled • jane@example.com",
        element("AXButton", AXDescription="Send", AXHelp="Send the message"),
        AXSubrole="AXStandardWindow",
    )
    rig = app_rig(compose, main)
    waiter = fake_waiter(rig.clock)
    events = RecordingEvents()
    outlook = AxOutlook(
        rig.app, WindowLocator(rig.app), PopupMenus(rig.app, events, waiter), events
    )
    current = OutlookScreen(rig, outlook, main, title_bar, toggle, rows, field, events)
    toggle.actions["AXPress"] = lambda: flip(toggle)
    perform.actions["AXPress"] = lambda: enter_search(current)
    calendar_button.actions["AXPress"] = lambda: retitle(main, "Calendar")
    return current


def flip(toggle: FakeElement) -> None:
    other = toggle.attributes["AXValue"] == 0
    toggle.attributes["AXValue"] = 1 if other else 0


def retitle(target: FakeElement, title: str) -> None:
    target.attributes["AXTitle"] = title


def enter_search(current: OutlookScreen) -> None:
    retitle(current.main, SEARCH_TITLE)
    cancel = element("AXButton", AXDescription="Cancel Search", AXHelp="Cancel")
    cancel.actions["AXPress"] = lambda: leave_search(current, cancel)
    current.title_bar.add(cancel)


def leave_search(current: OutlookScreen, cancel: FakeElement) -> None:
    retitle(current.main, INBOX_TITLE)
    current.field.attributes["AXValue"] = ""
    cancel.remove()


def test_main_window_is_found_by_its_module_switcher_not_its_title() -> None:
    current = screen()

    assert current.outlook.window_title() == INBOX_TITLE
    assert current.outlook.account() == "jane@example.com"
    assert current.outlook.mode() is OutlookMode.MAIL
    assert current.outlook.main_probe().handle is current.main


def test_rows_carry_descriptions_selection_and_levels() -> None:
    current = screen()

    messages = current.outlook.message_rows()
    folders = current.outlook.folder_rows()

    assert [node.selected for node in messages] == [True, False]
    assert messages[1].description.startswith("  Bob Sample")
    assert [(node.description, node.level) for node in folders] == [
        ("jane@example.com", 0),
        ("Inbox; 2 unread messages", 1),
    ]
    current.outlook.deselect(messages[0])
    assert current.rows[0].attributes["AXSelected"] is False


def test_focused_tab_is_read_and_switched_by_the_header_checkbox() -> None:
    current = screen()

    assert current.outlook.focused_tab() is FocusedTab.FOCUSED
    current.outlook.switch_tab(FocusedTab.OTHER)
    assert current.outlook.focused_tab() is FocusedTab.OTHER
    current.outlook.switch_tab(FocusedTab.OTHER)
    assert len(current.rig.backend.actions_of("AXPress")) == 1


def test_search_runs_reports_its_state_and_is_cancelled() -> None:
    current = screen()

    assert current.outlook.search_state() is None
    current.outlook.run_search("atlas", SearchScope.MAILBOX)

    assert current.field.attributes["AXValue"] == "atlas"
    assert current.outlook.mode() is OutlookMode.SEARCH
    assert current.outlook.search_state() == SearchState("atlas", SearchScope.MAILBOX)
    current.outlook.cancel_search()
    assert current.outlook.mode() is OutlookMode.MAIL


def test_module_switch_waits_for_the_new_mode() -> None:
    current = screen()

    current.outlook.show_module(OutlookModule.CALENDAR)

    assert current.outlook.mode() is OutlookMode.CALENDAR


def test_row_custom_actions_are_listed_and_performed_without_selecting() -> None:
    current = screen()
    row = current.outlook.message_rows()[1]

    assert current.outlook.row_actions(row) == ("Archive",)
    current.outlook.perform_row_action(row, "Archive")

    assert current.rig.backend.performed[-1] == (current.rows[1], ARCHIVE_ACTION)
    assert current.rig.backend.sets_of("AXSelected") == []
    with pytest.raises(ElementNotFoundError):
        current.outlook.perform_row_action(row, "Settings...")
    with pytest.raises(ForbiddenActionError):
        current.outlook.perform_row_action(row, "Report Junk")


def test_sheets_are_read_and_teaching_popovers_dismissed() -> None:
    current = screen()
    got_it = element("AXButton", AXTitle="Got it")
    popover = element("AXPopover", element("AXStaticText", AXValue="New: undo send"), got_it)
    got_it.actions["AXPress"] = popover.remove
    current.main.add(popover)
    current.main.add(
        element("AXSheet", element("AXButton", AXTitle="Cancel"), AXDescription="alert")
    )

    sheets = current.outlook.alert_sheets()

    assert [node.role for node in sheets] == ["AXSheet"]
    assert sheets[0].find(lambda node: node.label() == "Cancel") is not None
    assert current.outlook.dismiss_known_popovers() == 1
    assert current.outlook.dismiss_known_popovers() == 0


def test_actions_need_an_element_handle() -> None:
    current = screen()

    with pytest.raises(ElementNotFoundError):
        current.outlook.select(Node(role="AXRow"))


def test_search_waits_for_the_perform_button_and_fails_only_after_the_timeout() -> None:
    current = screen()
    perform = next(
        kid
        for kid in current.field.children
        if kid.attributes.get("AXDescription") == "Perform Search"
    )
    current.field.on_set["AXValue"] = lambda _value: perform.remove()
    started = current.rig.clock.now

    with pytest.raises(ElementNotFoundError, match="Perform Search"):
        current.outlook.run_search("atlas", SearchScope.MAILBOX)

    assert current.rig.clock.now - started >= current.rig.app.waiter.timeout


def test_a_scope_chosen_with_return_that_starts_the_search_is_not_searched_twice() -> None:
    current = screen()
    scope = next(
        kid for kid in current.field.children if kid.attributes.get("AXRole") == "AXPopUpButton"
    )
    wanted = element("AXCell", element("AXStaticText", AXValue="All Mailboxes"))
    table = element("AXTable", element("AXRow", wanted, AXSelected=False))
    scope_window = window("menuWindow", table, AXSubrole="AXUnknown")
    windows = current.rig.root.attributes["AXWindows"]
    assert isinstance(windows, list)

    def show_scopes() -> int:
        windows.append(scope_window)
        return 0

    def choose(_pid: int, shortcut: Shortcut) -> None:
        if shortcut == Shortcut(Key.RETURN):
            windows.remove(scope_window)
            scope.attributes["AXValue"] = "All Mailboxes"
            enter_search(current)

    scope.actions["AXPress"] = show_scopes
    current.events.on_press = choose

    current.outlook.run_search("atlas", SearchScope.ALL)

    assert current.outlook.mode() is OutlookMode.SEARCH
    assert current.rig.backend.actions_of("AXPress") == [scope]


def add_suggestions_window(current: OutlookScreen) -> FakeElement:
    popup = element("AXWindow", AXSubrole="AXUnknown", AXDescription=SUGGESTIONS_WINDOW)
    current.rig.root.add(popup)
    windows = current.rig.root.attributes["AXWindows"]
    assert isinstance(windows, list)
    windows.append(popup)
    return popup


def drop_window(current: OutlookScreen, popup: FakeElement) -> None:
    windows = current.rig.root.attributes["AXWindows"]
    assert isinstance(windows, list)
    windows.remove(popup)
    popup.remove()


def test_search_suggestions_are_closed_with_escape_only_when_shown() -> None:
    current = screen()

    assert current.outlook.close_search_suggestions() == 0
    assert current.events.shortcuts() == []

    popup = add_suggestions_window(current)
    current.events.on_press = lambda _pid, _shortcut: drop_window(current, popup)

    assert current.outlook.close_search_suggestions() == 1
    assert current.events.shortcuts() == [Shortcut(Key.ESCAPE)]


def test_search_suggestions_that_stay_open_raise_after_two_tries() -> None:
    current = screen()
    add_suggestions_window(current)

    with pytest.raises(CleanupError, match="suggestions"):
        current.outlook.close_search_suggestions()

    assert current.events.shortcuts() == [Shortcut(Key.ESCAPE), Shortcut(Key.ESCAPE)]


def search_without_cancel(current: OutlookScreen) -> None:
    current.outlook.run_search("atlas", SearchScope.MAILBOX)
    cancel = next(
        kid
        for kid in current.title_bar.children
        if kid.attributes.get("AXDescription") == "Cancel Search"
    )
    cancel.remove()


def test_search_without_a_cancel_button_is_left_with_escape_in_the_field() -> None:
    current = screen()
    search_without_cancel(current)
    current.events.on_press = lambda _pid, _shortcut: retitle(current.main, INBOX_TITLE)
    focus_before = len(current.rig.backend.sets_of("AXFocused"))

    current.outlook.cancel_search()

    assert current.events.shortcuts() == [Shortcut(Key.ESCAPE)]
    assert current.rig.backend.sets_of("AXFocused")[focus_before:] == [True]
    assert current.outlook.mode() is OutlookMode.MAIL


def test_an_escape_swallowed_by_the_suggestions_list_is_pressed_again() -> None:
    current = screen()
    search_without_cancel(current)
    presses: list[Shortcut] = []

    def swallow_first(_pid: int, shortcut: Shortcut) -> None:
        presses.append(shortcut)
        if len(presses) > 1:
            retitle(current.main, INBOX_TITLE)

    current.events.on_press = swallow_first

    current.outlook.cancel_search()

    assert current.events.shortcuts() == [Shortcut(Key.ESCAPE), Shortcut(Key.ESCAPE)]
    assert current.outlook.mode() is OutlookMode.MAIL


def test_search_that_never_ends_raises_after_every_escape() -> None:
    current = screen()
    search_without_cancel(current)

    with pytest.raises(UiTimeoutError, match="leave search"):
        current.outlook.cancel_search()

    assert current.events.shortcuts() == [Shortcut(Key.ESCAPE)] * 3


def test_cancel_search_outside_search_does_nothing() -> None:
    current = screen()

    current.outlook.cancel_search()

    assert current.events.shortcuts() == []
    assert current.rig.backend.actions_of("AXPress") == []


def test_a_new_search_leaves_the_running_one_first() -> None:
    current = screen()
    current.outlook.run_search("atlas", SearchScope.MAILBOX)

    current.outlook.run_search("budget", SearchScope.MAILBOX)

    assert current.outlook.search_state() == SearchState("budget", SearchScope.MAILBOX)
    pressed = [
        str(node.attributes.get("AXDescription")) for node, _action in current.rig.backend.performed
    ]
    assert pressed.count("Cancel Search") == 1


def test_search_focuses_the_field_so_its_buttons_are_exposed() -> None:
    current = screen()
    hidden = list(current.field.children)
    current.field.children.clear()

    def expose(value: object) -> None:
        current.field.attributes["AXFocused"] = value
        if value is True and not current.field.children:
            current.field.add(*hidden)

    current.field.on_set["AXFocused"] = expose

    current.outlook.run_search("atlas", SearchScope.MAILBOX)

    assert current.outlook.mode() is OutlookMode.SEARCH
    assert current.rig.backend.sets_of("AXFocused") == [True]


def test_collapse_focuses_the_message_list_and_presses_left() -> None:
    current = screen()

    current.outlook.collapse_selected()

    focused = current.rig.backend.sets[-1]
    assert focused[0].attributes.get("AXRole") == "AXTable"
    assert focused[1:] == ("AXFocused", True)
    assert current.events.shortcuts() == [Shortcut(Key.LEFT)]
