from __future__ import annotations

from fakes.ax_backend import FakeElement, element
from fakes.events import RecordingEvents
from fakes.p1_mail import MailGateway
from fakes.trees import app_rig, menu_bar, menu_item, outlook_main_window, window

from office_mcp.features.outlook.mail.search import AxSearchScope
from office_mcp.features.outlook.people.service import AxSuggestions
from office_mcp.features.outlook.rows import SearchScope
from office_mcp.features.outlook.service import MenuUnreadMarker, ToolbarUnreadMarker
from office_mcp.features.outlook.suggestions import SUGGESTIONS_WINDOW
from office_mcp.shared.events import Key, Shortcut
from office_mcp.shared.menus import MenuBar, PopupMenus
from office_mcp.shared.node import Node

SUGGESTION = "Best Match people suggestion 1 of 1 Jane Doe jane@example.com"


def scope_field() -> tuple[FakeElement, FakeElement]:
    popup = element("AXPopUpButton", AXDescription="Search Scope", AXValue="Current Mailbox")
    items = []
    for scope in SearchScope:
        cell = menu_item(scope.value)

        def pick(value: str = scope.value) -> None:
            popup.attributes["AXValue"] = value
            for child in list(popup.children):
                child.remove()

        cell.actions["AXPress"] = pick
        items.append(cell)

    def show() -> None:
        popup.add(element("AXMenu", *items))

    popup.actions["AXPress"] = show
    field = element("AXTextField", popup, AXDescription="Search")
    return field, popup


def test_menu_marker_presses_the_explicit_mark_as_unread_item() -> None:
    item = menu_item("Mark as Unread")
    rig = app_rig(
        outlook_main_window(),
        menu_bar=menu_bar(("Message", [menu_item("Toggle Unread"), item])),
    )

    MenuUnreadMarker(MenuBar(rig.app)).mark_unread()

    assert rig.backend.actions_of("AXPress") == [item]


class ToolbarGateway(MailGateway):
    def toolbar_button(self, name: str) -> Node | None:
        return Node(role="AXButton", description=name, handle=name)


def test_toolbar_marker_presses_the_mark_unread_button() -> None:
    gateway = ToolbarGateway()

    ToolbarUnreadMarker(gateway).mark_unread()

    assert gateway.calls == ["press_button:Mark Unread"]


def test_search_scope_reads_and_switches_through_the_scope_menu() -> None:
    field, popup = scope_field()
    main = outlook_main_window(element("AXGroup", field))
    rig = app_rig(main)
    control = AxSearchScope(
        rig.app,
        lambda: main,
        PopupMenus(rig.app, RecordingEvents(), rig.app.waiter),
        rig.app.waiter,
    )

    assert control.current() is SearchScope.MAILBOX
    control.choose(SearchScope.ALL)

    assert control.current() is SearchScope.ALL
    assert popup.children == []


def test_search_scope_focuses_the_field_when_its_buttons_are_hidden() -> None:
    field, _popup = scope_field()
    hidden = list(field.children)
    field.children.clear()

    def expose(value: object) -> None:
        field.attributes["AXFocused"] = value
        if value is True and not field.children:
            field.add(*hidden)

    field.on_set["AXFocused"] = expose
    main = outlook_main_window(element("AXGroup", field))
    rig = app_rig(main)
    control = AxSearchScope(
        rig.app,
        lambda: main,
        PopupMenus(rig.app, RecordingEvents(), rig.app.waiter),
        rig.app.waiter,
    )

    assert control.current() is SearchScope.MAILBOX
    assert rig.backend.sets_of("AXFocused") == [True]


def test_search_scope_is_unknown_without_a_field() -> None:
    main = outlook_main_window()
    rig = app_rig(main)
    control = AxSearchScope(
        rig.app,
        lambda: main,
        PopupMenus(rig.app, RecordingEvents(), rig.app.waiter),
        rig.app.waiter,
    )

    assert control.current() is None


def test_suggestions_type_read_and_clear() -> None:
    table = element(
        "AXTable",
        element("AXRow", element("AXCell", AXDescription=SUGGESTION)),
        element("AXRow", element("AXCell", AXDescription="Best Match text suggestion 2 of 2 x")),
    )
    field = element("AXTextField", table, AXDescription="Search", AXValue="")
    main = outlook_main_window(element("AXGroup", field))
    rig = app_rig(main)
    popup = window("", AXSubrole="AXUnknown", AXDescription=SUGGESTIONS_WINDOW)
    windows = rig.root.attributes["AXWindows"]
    assert isinstance(windows, list)

    def show_history(value: object) -> None:
        field.attributes["AXValue"] = value
        if value == "":
            windows.append(popup)

    def close_popup(_pid: int, _shortcut: Shortcut) -> None:
        windows.remove(popup)
        table.remove()

    field.on_set["AXValue"] = show_history
    events = RecordingEvents()
    events.on_press = close_popup
    source = AxSuggestions(rig.app, lambda: main, events, rig.app.waiter)

    source.type("jane")
    rows = source.rows()
    source.clear()

    assert rig.backend.sets_of("AXFocused") == [True]
    assert rig.backend.sets_of("AXValue") == ["jane", ""]
    assert rows == [SUGGESTION, "Best Match text suggestion 2 of 2 x"]
    assert events.shortcuts() == [Shortcut(Key.ESCAPE)]
    assert popup not in windows
    assert source.rows() == []


def test_clearing_without_a_suggestions_window_presses_nothing() -> None:
    field = element("AXTextField", AXDescription="Search", AXValue="")
    main = outlook_main_window(element("AXGroup", field))
    rig = app_rig(main)
    events = RecordingEvents()
    source = AxSuggestions(rig.app, lambda: main, events, rig.app.waiter)

    source.type("jane")
    source.clear()

    assert events.shortcuts() == []
