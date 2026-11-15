from __future__ import annotations

from dataclasses import dataclass

import pytest
from fakes.ax_backend import FakeElement, element
from fakes.events import RecordingEvents
from fakes.trees import AppRig, app_rig, menu_bar, menu_item, outlook_main_window

from office_mcp.features.outlook.actions.ax import AxActions, parse_target
from office_mcp.features.outlook.actions.gateway import MoveSheet, MoveTarget
from office_mcp.shared.errors import (
    ActionNotVerifiedError,
    ActivationError,
    AmbiguousTargetError,
)
from office_mcp.shared.events import Key, Shortcut
from office_mcp.shared.geometry import POINT, decode_pair
from office_mcp.shared.menus import MenuBar, PopupMenus
from office_mcp.shared.windows import WindowLocator

ACCOUNT = "jane@example.com"


@dataclass
class SheetRig:
    rig: AppRig
    actions: AxActions
    events: RecordingEvents
    sheet: FakeElement
    box: FakeElement
    area: FakeElement
    travel: FakeElement


def folder_group(name: str, y: float) -> FakeElement:
    return element(
        "AXGroup", AXDescription=f"{name}, {ACCOUNT}", AXPosition=(40.0, y), AXSize=(200.0, 20.0)
    )


def build(*, frontmost: bool = True) -> SheetRig:
    travel = folder_group("Travel", 110.0)
    area = element(
        "AXScrollArea",
        folder_group("Inbox", 90.0),
        travel,
        AXPosition=(40.0, 80.0),
        AXSize=(200.0, 100.0),
    )
    box = element("AXTextField", AXDescription="search box", AXValue="")
    sheet = element(
        "AXSheet",
        box,
        area,
        element("AXButton", AXTitle="Cancel"),
        element("AXButton", AXTitle="Move", AXEnabled=True),
    )
    table = element("AXTable", AXDescription="Message List")
    main = outlook_main_window(table, sheet)
    bar = menu_bar(
        (
            "Message",
            [
                menu_item(
                    "Move",
                    menu_item("Travel"),
                    menu_item("Choose Folder...", char="M", modifiers=1),
                    menu_item("Copy to Folder...", char="C", modifiers=1),
                )
            ],
        )
    )
    rig = app_rig(main, menu_bar=bar)
    rig.root.attributes["AXFrontmost"] = frontmost
    events = RecordingEvents()
    app = rig.app
    actions = AxActions(
        app,
        WindowLocator(app),
        MenuBar(app),
        PopupMenus(app, events, app.waiter),
        events,
        app.waiter,
    )
    return SheetRig(rig, actions, events, sheet, box, area, travel)


def test_parse_target_splits_folder_and_account() -> None:
    assert parse_target("Clients, Projects, jane@example.com") == MoveTarget(
        "Clients, Projects", ACCOUNT, "Clients, Projects, jane@example.com"
    )
    assert parse_target("Inbox") is None


def test_decode_pair_reads_tuples_and_described_values() -> None:
    described = "<AXValue> {value = x:12.500000 y:-4.000000 type = kAXValueCGPointType}"

    assert decode_pair((1, 2), POINT, ("x", "y")) == (1.0, 2.0)
    assert decode_pair(described, POINT, ("x", "y")) == (12.5, -4.0)
    assert decode_pair(None, POINT, ("x", "y")) is None


def test_menu_titles_lists_the_move_submenu() -> None:
    setup = build()

    titles = setup.actions.menu_titles(("Message", "Move"))

    assert titles == ("Travel", "Choose Folder...", "Copy to Folder...")


def test_open_row_menu_shows_the_message_list_menu() -> None:
    setup = build()
    table = setup.rig.root.children[0].children[1]

    def show_menu() -> None:
        table.add(element("AXMenu"))

    table.actions["AXShowMenu"] = show_menu

    menu = setup.actions.open_row_menu()

    assert menu.role == "AXMenu"
    assert setup.rig.backend.actions_of("AXShowMenu") == [table]


def test_filter_sets_the_search_box_and_lists_targets() -> None:
    setup = build()
    sheet = MoveSheet(setup.sheet)

    targets = setup.actions.filter_move_sheet(sheet, "Travel")

    assert setup.box.attributes["AXValue"] == "Travel"
    assert [target.folder for target in targets] == ["Inbox", "Travel"]
    assert setup.actions.sheet_open(sheet) is True


def test_pick_refuses_a_folder_below_the_first_entry() -> None:
    setup = build()
    sheet = MoveSheet(setup.sheet)

    with pytest.raises(AmbiguousTargetError, match="nothing was moved") as caught:
        setup.actions.pick_move_target(sheet, parse_target(f"Travel, {ACCOUNT}") or pytest.fail())

    assert caught.value.candidates == (f"Inbox, {ACCOUNT}",)
    assert setup.events.presses == []


def test_pick_of_the_first_folder_only_returns_to_the_top() -> None:
    setup = build()

    setup.actions.pick_move_target(
        MoveSheet(setup.sheet), parse_target(f"Inbox, {ACCOUNT}") or pytest.fail()
    )

    assert setup.events.shortcuts() == [Shortcut(Key.UP), Shortcut(Key.UP)]
    assert setup.box.attributes["AXFocused"] is True
    assert setup.events.clicks == []


def test_pick_refuses_a_folder_the_sheet_no_longer_lists() -> None:
    setup = build()
    setup.travel.remove()

    with pytest.raises(ActionNotVerifiedError, match="changed while picking"):
        setup.actions.pick_move_target(
            MoveSheet(setup.sheet), parse_target(f"Travel, {ACCOUNT}") or pytest.fail()
        )

    assert setup.events.presses == []


def test_pick_refuses_when_outlook_is_not_frontmost() -> None:
    setup = build(frontmost=False)

    with pytest.raises(ActivationError):
        setup.actions.pick_move_target(
            MoveSheet(setup.sheet), parse_target(f"Travel, {ACCOUNT}") or pytest.fail()
        )

    assert setup.events.presses == []


def test_cancel_presses_the_sheet_cancel_button() -> None:
    setup = build()
    cancel = setup.sheet.children[2]
    cancel.actions["AXPress"] = setup.sheet.remove

    setup.actions.cancel_move(MoveSheet(setup.sheet))

    assert setup.actions.sheet_open(MoveSheet(setup.sheet)) is False
