from __future__ import annotations

import pytest
from fakes.p2_outlook import actions_rig, atlas, default_folders, offsite

from office_mcp.features.outlook.actions.gateway import MoveTarget
from office_mcp.features.outlook.actions.move import (
    FolderDestination,
    TabDestination,
    recent_items,
    resolve_destination,
    same_folder,
    sheet_matches,
)
from office_mcp.features.outlook.listing import parse_folders
from office_mcp.features.outlook.window import FocusedTab
from office_mcp.shared.errors import AmbiguousTargetError, TargetError, TargetNotFoundError
from office_mcp.shared.events import Key, Modifier, Shortcut

ACCOUNT = "jane@example.com"
FOLDERS = parse_folders(default_folders())


def target(folder: str, account: str = ACCOUNT) -> MoveTarget:
    return MoveTarget(folder=folder, account=account, description=f"{folder}, {account}")


def test_destination_matches_name_or_path_and_skips_favorites() -> None:
    by_name = resolve_destination("travel", FOLDERS)
    by_path = resolve_destination("jane@example.com/Clients/Projects", FOLDERS)
    inbox = resolve_destination("Inbox", FOLDERS)

    assert by_name == FolderDestination("Travel", "jane@example.com/Travel", unique_name=True)
    assert by_path == FolderDestination(
        "Projects", "jane@example.com/Clients/Projects", unique_name=False
    )
    assert inbox == FolderDestination("Inbox", "jane@example.com/Inbox", unique_name=True)


def test_destination_keywords_pick_the_focused_and_other_tabs() -> None:
    assert resolve_destination("Other", FOLDERS) == TabDestination(FocusedTab.OTHER)
    assert resolve_destination(" focused ", FOLDERS) == TabDestination(FocusedTab.FOCUSED)


def test_shared_folder_names_need_a_path() -> None:
    with pytest.raises(AmbiguousTargetError) as caught:
        resolve_destination("Projects", FOLDERS)

    assert caught.value.candidates == (
        "jane@example.com/Projects",
        "jane@example.com/Clients/Projects",
    )


def test_unknown_or_empty_destination_is_refused() -> None:
    with pytest.raises(TargetNotFoundError, match="Folders:"):
        resolve_destination("Receipts", FOLDERS)
    with pytest.raises(TargetError):
        resolve_destination("  ", FOLDERS)


def test_same_folder_compares_names_and_paths() -> None:
    travel = FolderDestination("Travel", "jane@example.com/Travel", unique_name=True)

    assert same_folder("Travel", travel)
    assert same_folder("jane@example.com/travel", travel)
    assert not same_folder("Inbox", travel)
    assert not same_folder("", travel)


def test_sheet_matches_need_the_exact_folder_in_the_current_account() -> None:
    targets = [
        target("Travel"),
        target("Travel", "On My Computer"),
        target("Travel", "bob@example.com"),
        target("Travel plans"),
    ]

    assert sheet_matches(targets, "travel", ACCOUNT) == [target("Travel")]
    assert len(sheet_matches(targets, "Travel", None)) == 2


def test_recent_items_ignore_the_sheet_entries() -> None:
    titles = ("Travel", "Choose Folder...", "Copy to Folder...", "Travel plans")

    assert recent_items(titles, "travel") == ["Travel"]
    assert recent_items(titles, "Choose Folder...") == []


def test_move_uses_a_recent_folder_item_when_the_name_is_unique() -> None:
    rig = actions_rig({"m1": atlas()})
    rig.actions.move_titles = ("Travel", "Choose Folder...", "Copy to Folder...")

    record = rig.service.move([rig.key("m1")], "Travel", None, rig.deadline()).to_record()

    assert record["destination"] == "jane@example.com/Travel"
    assert record["results"][0]["verified"] is True
    assert "open_sheet" not in rig.outlook.calls
    assert "menu:Message > Move > Travel" in rig.outlook.calls


def test_move_through_the_sheet_picks_once_and_confirms() -> None:
    rig = actions_rig({"m1": atlas(unread=True)})
    rig.actions.targets = (target("Travel"), target("Travel", "On My Computer"))

    result = rig.service.move([rig.key("m1")], "Travel", None, rig.deadline())

    assert result.items[0].detail == "moved to jane@example.com/Travel with the Move sheet"
    assert rig.visible() == [
        "select:m1",
        "menu_titles:Message",
        "menu:Message > Mark as Unread",
        "menu_titles:Message > Move",
        "open_sheet",
        "filter:Travel",
        "pick:Travel, jane@example.com",
        "confirm_move",
    ]


def test_move_ambiguity_cancels_the_sheet_and_moves_nothing() -> None:
    rig = actions_rig({"m1": atlas()})
    rig.actions.targets = (target("Travel"), target("Travel"))

    result = rig.service.move([rig.key("m1")], "Travel", None, rig.deadline())

    assert result.items[0].ok is False
    assert "2 folders named 'Travel'" in result.items[0].detail
    assert "cancel_sheet" in rig.outlook.calls
    assert "confirm_move" not in rig.outlook.calls
    assert "m1" in rig.outlook.mails


def test_move_cancels_when_the_pick_fails() -> None:
    rig = actions_rig({"m1": atlas()})
    rig.actions.targets = (target("Travel"),)
    rig.actions.pick_fails = True

    result = rig.service.move([rig.key("m1")], "Travel", None, rig.deadline())

    assert "changed while picking 'Travel'" in result.items[0].detail
    assert rig.outlook.calls.count("cancel_sheet") == 1
    assert "confirm_move" not in rig.outlook.calls


def test_move_refuses_a_folder_that_is_not_first_in_the_sheet() -> None:
    rig = actions_rig({"m1": atlas()})
    rig.actions.targets = (target("Travel plans"), target("Travel"))

    result = rig.service.move([rig.key("m1")], "Travel", None, rig.deadline())

    assert (result.items[0].ok, result.items[0].verified) == (False, False)
    assert "nothing was moved" in result.items[0].detail
    assert f"Candidates: Travel plans, {ACCOUNT}" in result.items[0].detail
    assert rig.outlook.calls.count("cancel_sheet") == 1
    assert not any(call.startswith("pick:") for call in rig.outlook.calls)
    assert "confirm_move" not in rig.outlook.calls
    assert "m1" in rig.outlook.mails


def test_move_cancels_a_sheet_that_stays_open_before_the_next_item() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.actions.targets = (target("Travel"),)
    rig.actions.sheet_sticks = True

    result = rig.service.move([rig.key("m1"), rig.key("m2")], "Travel", None, rig.deadline())

    assert "did not close" in result.items[0].detail
    calls = rig.visible()
    first_cancel = calls.index("cancel_sheet")
    assert calls.index("confirm_move") < first_cancel < calls.index("open_sheet", first_cancel)
    assert rig.actions.sheet is False


def test_move_never_moves_again_after_a_failed_check() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.actions.targets = (target("Travel"),)
    rig.actions.move_works = False

    result = rig.service.move([rig.key("m1")], "Travel", None, rig.deadline())

    assert (result.items[0].ok, result.items[0].verified) == (False, False)
    assert "still listed" in result.items[0].detail
    assert rig.outlook.calls.count("confirm_move") == 1
    assert rig.outlook.calls.count("open_sheet") == 1


def test_move_refuses_the_folder_it_is_already_in() -> None:
    rig = actions_rig({"m1": atlas()})

    result = rig.service.move([rig.key("m1")], "Inbox", None, rig.deadline())

    assert result.items[0].ok is False
    assert "already in" in result.items[0].detail
    assert not any(call.startswith("select:") for call in rig.outlook.calls)


def test_move_to_other_uses_the_message_menu_item() -> None:
    rig = actions_rig({"m1": atlas()})

    record = rig.service.move([rig.key("m1")], "other", None, rig.deadline()).to_record()

    assert record["destination"] == "other"
    assert record["results"][0]["verified"] is True
    assert rig.actions.presses == [(("Message", "Move to Other"), Shortcut(Key.O, Modifier.OPTION))]


def test_move_to_the_tab_it_is_in_is_refused() -> None:
    rig = actions_rig({"m1": atlas()})

    result = rig.service.move([rig.key("m1")], "focused", None, rig.deadline())

    assert "already in the focused tab" in result.items[0].detail
    assert rig.actions.presses == []


def test_unknown_destination_stops_before_any_item() -> None:
    rig = actions_rig({"m1": atlas()})

    with pytest.raises(TargetNotFoundError):
        rig.service.move([rig.key("m1")], "Receipts", None, rig.deadline())

    assert rig.scope_log == ["enter foreground=True", "exit"]
    assert not any(call.startswith("select:") for call in rig.outlook.calls)
