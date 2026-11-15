from __future__ import annotations

import pytest
from fakes.p2_outlook import (
    ARCHIVE,
    DELETE,
    MARK_READ,
    MARK_UNREAD,
    actions_rig,
    atlas,
    deleted_items_alert,
    lunch,
    offsite,
)

from office_mcp.features.outlook.actions.service import validated_keys
from office_mcp.features.outlook.rows import SearchScope, SearchState
from office_mcp.features.outlook.window import OutlookMode
from office_mcp.shared.errors import ForbiddenActionError, TargetError
from office_mcp.shared.events import Key, Modifier, Shortcut
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.node import Node
from office_mcp.shared.runner import Deadline


class Ticks:
    def __init__(self) -> None:
        self.now = -1.0

    def __call__(self) -> float:
        self.now += 1.0
        return self.now


@pytest.mark.parametrize("refs", [["3"], ["Project Atlas"], ["ABCDEF123456"], []])
def test_non_key_refs_are_rejected_before_any_ui_call(refs: list[str]) -> None:
    rig = actions_rig({"m1": atlas()})

    with pytest.raises(TargetError):
        rig.service.archive(refs, None, rig.deadline())

    assert rig.outlook.calls == []
    assert rig.scope_log == []


def test_more_than_twenty_five_keys_are_refused() -> None:
    keys = [f"{index:012x}" for index in range(26)]

    with pytest.raises(TargetError, match="at most 25"):
        validated_keys(keys)


def test_repeated_keys_are_acted_on_once() -> None:
    assert validated_keys(["0123456789ab", "0123456789ab "]) == ("0123456789ab",)


def test_vanished_key_fails_that_item_without_acting() -> None:
    rig = actions_rig({"m1": atlas()})

    result = rig.service.delete(["0123456789ab"], None, rig.deadline())

    record = result.to_record()
    assert record["failed"] == 1
    assert "changed or a day passed" in record["results"][0]["detail"]
    assert not any(call.startswith(("menu:", "select:")) for call in rig.outlook.calls)
    assert rig.scope_log == ["enter foreground=True", "ensure_front", "exit"]


def test_duplicated_key_fails_that_item_without_acting() -> None:
    rig = actions_rig({"m1": atlas(), "m2": atlas()})
    key = rig.key("m1")

    result = rig.service.archive([key], None, rig.deadline())

    assert result.items[0].ok is False
    assert "matches several conversations" in result.items[0].detail
    assert not any(call.startswith(("row_action", "menu:", "select:")) for call in rig.visible())


def test_archive_uses_the_row_action_and_never_selects() -> None:
    rig = actions_rig({"m1": atlas(unread=True), "m2": offsite()})
    rig.outlook.actions = {"m1": ("Settings...", "Archive")}
    key = rig.key("m1")

    result = rig.service.archive([key], None, rig.deadline())

    item = result.items[0]
    assert (item.ok, item.verified, item.location) == (True, True, "Archive")
    assert item.detail == "archived with the row action"
    assert rig.visible() == ["row_action:m1:Archive"]
    assert result.focus is FocusOutcome.RESTORED


def test_archive_falls_back_to_the_menu_and_keeps_the_unread_state() -> None:
    rig = actions_rig({"m1": atlas(unread=True)})
    key = rig.key("m1")

    result = rig.service.archive([key], None, rig.deadline())

    assert result.items[0].verified is True
    assert rig.visible() == [
        "select:m1",
        "menu_titles:Message",
        "menu:Message > Mark as Unread",
        "menu:Message > Archive",
    ]
    assert rig.actions.presses[-1] == (ARCHIVE, Shortcut(Key.E, Modifier.CONTROL))


def test_archive_row_action_that_seems_to_do_nothing_is_never_repeated() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.outlook.actions = {"m1": ("Archive",)}
    rig.outlook.row_archive_works = False
    key = rig.key("m1")

    result = rig.service.archive([key], None, rig.deadline())

    item = result.items[0]
    assert (item.ok, item.verified) == (False, False)
    assert "check Archive" in item.detail
    assert rig.visible() == ["row_action:m1:Archive"]
    assert rig.actions.presses == []
    assert set(rig.outlook.mails) == {"m1", "m2"}


def test_delete_presses_the_exact_delete_item_and_reports_the_message_count() -> None:
    rig = actions_rig({"m1": atlas(count=3)})
    key = rig.key("m1")

    record = rig.service.delete([key], None, rig.deadline()).to_record()

    assert record["moved_to"] == "Deleted Items"
    assert record["results"][0]["detail"] == "3 messages moved to Deleted Items"
    assert rig.actions.presses == [(DELETE, Shortcut(Key.DELETE, Modifier.COMMAND))]


def test_delete_is_refused_inside_deleted_items() -> None:
    rig = actions_rig({"m1": atlas()}, folder="Deleted Items")
    key = rig.key("m1")

    with pytest.raises(ForbiddenActionError, match="permanent"):
        rig.service.delete([key], None, rig.deadline())

    assert rig.actions.presses == []
    assert rig.scope_log[-1] == "exit"


def test_delete_is_refused_in_a_subfolder_of_deleted_items() -> None:
    rig = actions_rig({"m1": atlas()}, folder="Old")
    key = rig.key("m1")

    with pytest.raises(ForbiddenActionError):
        rig.service.delete([key], None, rig.deadline())

    assert rig.actions.presses == []


def test_delete_is_refused_for_a_search_result_in_deleted_items() -> None:
    rig = actions_rig({"m1": atlas(folder="Deleted Items")})
    rig.outlook.mode_value = OutlookMode.SEARCH
    rig.outlook.search = SearchState("atlas", SearchScope.MAILBOX)
    key = rig.key("m1")

    with pytest.raises(ForbiddenActionError):
        rig.service.delete([key], None, rig.deadline())

    assert rig.actions.presses == []
    assert not any(call.startswith("select:") for call in rig.outlook.calls)


def test_delete_is_refused_for_a_search_result_without_a_folder() -> None:
    rig = actions_rig({"m1": atlas()})
    rig.outlook.mode_value = OutlookMode.SEARCH
    rig.outlook.search = SearchState("atlas", SearchScope.MAILBOX)
    key = rig.key("m1")

    with pytest.raises(ForbiddenActionError, match="which folder"):
        rig.service.delete([key], None, rig.deadline())


def test_permanent_alert_is_cancelled_and_stops_the_batch_with_done_items() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.actions.alert = deleted_items_alert()
    rig.actions.deletes_before_alert = 1
    first, second = rig.key("m1"), rig.key("m2")

    with pytest.raises(ForbiddenActionError, match="permanently") as caught:
        rig.service.delete([first, second], None, rig.deadline())

    assert f"done before the stop: {first}" in (caught.value.__notes__ or [])
    assert "press_button:Cancel" in rig.outlook.calls
    assert rig.outlook.sheets == []
    assert "m2" in rig.outlook.mails
    assert rig.scope_log[-1] == "exit"


def test_neighbor_marked_read_by_outlook_is_marked_unread_again() -> None:
    rig = actions_rig({"m1": lunch(unread=True), "m2": atlas(), "m3": offsite(unread=True)})
    rig.outlook.auto_select_next = True
    key = rig.key("m2")

    result = rig.service.delete([key], None, rig.deadline())

    assert rig.outlook.mails["m3"].unread is True
    assert rig.outlook.mails["m1"].unread is True
    assert "marked the neighboring 'Bob Sample, Team offsite' unread again" in (
        result.items[0].detail
    )
    tail = rig.visible()[-4:]
    assert tail == [
        "select:m3",
        "menu_titles:Message",
        "menu:Message > Mark as Unread",
        "deselect:m3",
    ]


def test_deadline_stops_the_batch_and_lists_skipped_keys() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    first, second = rig.key("m1"), rig.key("m2")
    rig.outlook.actions = {"m1": ("Archive",), "m2": ("Archive",)}

    result = rig.service.archive([first, second], None, Deadline(Ticks(), 1.0))

    record = result.to_record()
    assert record["succeeded"] == 1
    assert record["skipped"] == [second]
    assert "m2" in rig.outlook.mails


def test_unread_mail_still_shown_unread_is_deleted_without_pressing_mark_as_unread() -> None:
    rig = actions_rig({"m1": atlas(unread=True), "m2": offsite()})
    rig.outlook.select_marks_read = False
    rig.actions.hides_offered_items = True

    result = rig.service.delete([rig.key("m1")], None, rig.deadline())

    assert result.items[0].ok is True
    assert "could not be kept unread" not in result.items[0].detail
    assert (MARK_UNREAD, None) not in rig.actions.presses
    assert set(rig.outlook.mails) == {"m2"}


def test_delete_with_a_deleted_items_hint_is_refused_before_any_row_is_touched() -> None:
    rig = actions_rig({"m1": atlas()})

    with pytest.raises(ForbiddenActionError, match="Delete is permanent"):
        rig.service.delete([rig.key("m1")], "Deleted Items", rig.deadline())

    assert not any(call.startswith("select:") for call in rig.outlook.calls)
    assert "m1" in rig.outlook.mails


def test_mark_read_selects_verifies_and_deselects() -> None:
    rig = actions_rig({"m1": atlas(unread=True)})
    key = rig.key("m1")

    result = rig.service.mark([key], True, None, rig.deadline())

    assert result.action == "mark_read"
    assert (result.items[0].verified, result.items[0].detail) == (True, "marked read")
    assert rig.visible() == ["select:m1", "deselect:m1"]


def test_mark_read_uses_the_menu_when_selecting_does_not_mark_it_read() -> None:
    rig = actions_rig({"m1": atlas(unread=True)})
    rig.outlook.select_marks_read = False
    key = rig.key("m1")

    result = rig.service.mark([key], True, None, rig.deadline())

    assert result.items[0].verified is True
    assert (MARK_READ, None) in rig.actions.presses


def test_mark_unread_presses_mark_as_unread_and_verifies() -> None:
    rig = actions_rig({"m1": atlas()})
    key = rig.key("m1")

    result = rig.service.mark([key], False, None, rig.deadline())

    assert result.items[0].detail == "marked unread"
    assert rig.actions.presses == [(MARK_UNREAD, None)]
    assert rig.outlook.mails["m1"].unread is True


def test_mark_leaves_matching_rows_alone() -> None:
    rig = actions_rig({"m1": atlas(unread=True)})
    key = rig.key("m1")

    result = rig.service.mark([key], False, None, rig.deadline())

    assert result.items[0].detail == "already unread"
    assert rig.visible() == []


def test_mark_unread_corrects_once_then_reports_failure() -> None:
    rig = actions_rig({"m1": atlas()})
    rig.actions.unread_sticks = False
    key = rig.key("m1")

    result = rig.service.mark([key], False, None, rig.deadline())

    assert result.items[0].ok is False
    assert "still shows it as read" in result.items[0].detail
    assert rig.actions.presses == [(MARK_UNREAD, None), (MARK_UNREAD, None)]


def test_scope_exit_is_recorded_after_every_item() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.outlook.actions = {"m1": ("Archive",), "m2": ("Archive",)}

    rig.service.archive([rig.key("m1"), rig.key("m2")], None, rig.deadline())

    assert rig.scope_log == ["enter foreground=True", "ensure_front", "ensure_front", "exit"]


def test_nothing_is_pressed_when_outlook_selects_a_different_row() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.outlook.selected = {"m2"}
    rig.outlook.select_ignored = True

    result = rig.service.delete([rig.key("m1")], None, rig.deadline())

    assert result.items[0].ok is False
    assert "did not select only" in result.items[0].detail
    assert rig.actions.presses == []
    assert set(rig.outlook.mails) == {"m1", "m2"}


def test_a_conversation_whose_first_select_lands_on_a_message_is_selected_again() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.outlook.first_select_lands_on_child = {"m1"}

    result = rig.service.delete([rig.key("m1")], None, rig.deadline())

    assert result.items[0].ok is True
    assert "child_select:m1" in rig.outlook.calls
    assert set(rig.outlook.mails) == {"m2"}


def test_other_selected_rows_are_cleared_before_delete() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.outlook.selected = {"m2"}
    rig.outlook.select_extends = True

    result = rig.service.delete([rig.key("m1")], None, rig.deadline())

    assert result.items[0].ok is True
    assert "deselect:m2" in rig.outlook.calls
    assert set(rig.outlook.mails) == {"m2"}


def test_delete_is_refused_while_the_pane_shows_a_multi_selection() -> None:
    rig = actions_rig({"m1": atlas(), "m2": offsite()})
    rig.outlook.hidden_selected = {"scrolled-away"}
    rig.outlook.select_extends = True
    rig.outlook.multi_pane = Node(
        role="AXGroup", children=(Node(role="AXStaticText", value="2 items selected"),)
    )

    result = rig.service.delete([rig.key("m1")], None, rig.deadline())

    assert result.items[0].ok is False
    assert "did not select only" in result.items[0].detail
    assert rig.actions.presses == []
    assert set(rig.outlook.mails) == {"m1", "m2"}
