from __future__ import annotations

import pytest
from fakes.p2_outlook import actions_rig, atlas

from office_mcp.features.outlook.actions.flags import FlagState, checked_verdict, menu_path


@pytest.mark.parametrize(
    ("state", "item"),
    [
        (FlagState.TODAY, "Today"),
        (FlagState.TOMORROW, "Tomorrow"),
        (FlagState.THIS_WEEK, "This Week"),
        (FlagState.NEXT_WEEK, "Next Week"),
        (FlagState.NO_DATE, "No Date"),
        (FlagState.CLEAR, "Clear Flag"),
    ],
)
def test_flag_states_map_to_menu_items(state: FlagState, item: str) -> None:
    assert menu_path(state) == ("Flag", item)


@pytest.mark.parametrize(
    "state", [FlagState.TODAY, FlagState.TOMORROW, FlagState.NO_DATE, FlagState.NEXT_WEEK]
)
def test_flag_chooses_the_item_from_the_row_context_menu(state: FlagState) -> None:
    rig = actions_rig({"m1": atlas()})

    result = rig.service.flag([rig.key("m1")], state, None, rig.deadline())

    item = result.items[0]
    assert (item.ok, item.verified) == (True, True)
    assert rig.visible() == [
        "select:m1",
        "row_menu",
        f"choose:{' > '.join(menu_path(state))}",
        "close_menu",
        "deselect:m1",
    ]
    assert rig.outlook.mails["m1"].flagged is True


def test_clear_presses_the_toggle_only_when_flagged() -> None:
    flagged = actions_rig({"m1": atlas(flagged=True)})
    plain = actions_rig({"m1": atlas()})

    cleared = flagged.service.flag([flagged.key("m1")], FlagState.CLEAR, None, flagged.deadline())
    untouched = plain.service.flag([plain.key("m1")], FlagState.CLEAR, None, plain.deadline())

    assert cleared.items[0].verified is True
    assert flagged.outlook.mails["m1"].flagged is False
    assert "choose:Flag > Clear Flag" in flagged.outlook.calls
    assert untouched.items[0].detail == "already not flagged"
    assert plain.visible() == []


def test_flag_marks_an_unread_conversation_unread_again() -> None:
    rig = actions_rig({"m1": atlas(unread=True)})

    rig.service.flag([rig.key("m1")], FlagState.TODAY, None, rig.deadline())

    assert rig.outlook.mails["m1"].unread is True
    assert rig.visible()[-2:] == ["menu:Message > Mark as Unread", "deselect:m1"]


def test_flag_falls_back_to_the_menu_check_mark() -> None:
    rig = actions_rig({"m1": atlas()})
    rig.actions.flag_works = False
    rig.actions.flag_checks = {"Today": True}

    result = rig.service.flag([rig.key("m1")], FlagState.TODAY, None, rig.deadline())

    assert result.items[0].verified is True
    assert rig.outlook.calls.count("row_menu") == 2
    assert rig.outlook.calls.count("close_menu") == 2


def test_flag_without_any_readable_state_is_not_verified() -> None:
    rig = actions_rig({"m1": atlas()})
    rig.actions.flag_works = False

    result = rig.service.flag([rig.key("m1")], FlagState.TODAY, None, rig.deadline())

    item = result.items[0]
    assert (item.ok, item.verified) == (True, False)
    assert "could not be read back" in item.detail


def test_checked_verdict_for_clear_needs_every_due_item_unchecked() -> None:
    unchecked = dict.fromkeys(("Today", "Tomorrow", "This Week", "Next Week", "No Date"), False)

    assert checked_verdict(FlagState.CLEAR, unchecked) is True
    assert checked_verdict(FlagState.CLEAR, {**unchecked, "Tomorrow": True}) is False
    assert checked_verdict(FlagState.CLEAR, {}) is None
    assert checked_verdict(FlagState.TODAY, {"Today": False}) is False


def test_changing_the_due_date_of_a_flagged_row_reads_the_menu_check() -> None:
    rig = actions_rig({"m1": atlas(flagged=True)})
    rig.actions.flag_works = False
    rig.actions.flag_checks = {"Today": True, "Next Week": False}

    result = rig.service.flag([rig.key("m1")], FlagState.NEXT_WEEK, None, rig.deadline())

    item = result.items[0]
    assert (item.ok, item.verified) == (False, False)
    assert "does not show the flag change" in item.detail


def test_changing_the_due_date_is_verified_by_the_new_check_mark() -> None:
    rig = actions_rig({"m1": atlas(flagged=True)})
    rig.actions.flag_checks = {"Today": False, "Next Week": True}

    result = rig.service.flag([rig.key("m1")], FlagState.NEXT_WEEK, None, rig.deadline())

    assert result.items[0].verified is True
    assert rig.outlook.calls.count("row_menu") == 2
