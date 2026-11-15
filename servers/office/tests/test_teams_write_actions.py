from __future__ import annotations

import pytest
from fakes.p5_rig import (
    ALICE,
    BOB,
    BOB_QUESTION,
    MY_ANSWER,
    SELF_NOTE,
    SELF_THREAD,
    standard_world,
    writing_rig,
)
from fakes.p5_world import CONFIRM_HANDLE, WorldOptions

from office_mcp.features.teams.writing.reactions import Reaction
from office_mcp.shared.errors import (
    ActionNotVerifiedError,
    ForbiddenActionError,
    RecipientNotAllowedError,
)


def toggle_presses(pressed: list[object]) -> list[object]:
    return [handle for handle in pressed if isinstance(handle, tuple) and handle[0] == "toggle"]


def test_react_turns_the_reaction_on_and_verifies_it() -> None:
    rig = writing_rig(standard_world())

    record = rig.service.react(
        BOB, BOB_QUESTION, Reaction.LIKE, remove=False, deadline=rig.deadline()
    ).to_record()

    assert record == {
        "ok": True,
        "message_id": BOB_QUESTION,
        "reaction": "like",
        "state": "on",
        "changed": True,
        "verified": True,
        "chat": "Bob Sample",
        "thread_id": BOB,
        "focus_restored": "restored",
        "notes": [],
    }
    assert rig.world.chat(BOB).messages[0].reactions == {"Like"}
    assert toggle_presses(rig.world.pressed) == [("toggle", BOB_QUESTION, "Like")]
    assert rig.world.open_thread == ALICE


def test_react_is_a_no_op_when_already_in_the_wanted_state() -> None:
    world = standard_world()
    world.chat(BOB).messages[0].reactions.add("Heart")
    rig = writing_rig(world)

    result = rig.service.react(BOB, BOB_QUESTION, Reaction.HEART, False, rig.deadline())

    assert result.change.changed is False
    assert result.change.state is True
    assert toggle_presses(world.pressed) == []


def test_react_removes_an_existing_reaction() -> None:
    world = standard_world()
    world.chat(BOB).messages[0].reactions.add("Laugh")
    rig = writing_rig(world)

    result = rig.service.react(BOB, BOB_QUESTION, Reaction.LAUGH, True, rig.deadline())

    assert result.to_record()["state"] == "off"
    assert result.change.verified is True
    assert world.chat(BOB).messages[0].reactions == set()


def test_react_falls_back_to_the_message_menu_and_still_verifies() -> None:
    rig = writing_rig(standard_world(WorldOptions(popover=False)))

    result = rig.service.react(BOB, BOB_QUESTION, Reaction.SURPRISED, False, rig.deadline())

    assert result.change.verified is True
    assert BOB_QUESTION in rig.world.menus_opened
    assert len(toggle_presses(rig.world.pressed)) == 1


def test_react_refuses_to_press_when_the_state_cannot_be_read() -> None:
    rig = writing_rig(standard_world(WorldOptions(reaction_state_visible=False)))

    with pytest.raises(ActionNotVerifiedError, match="could not be read"):
        rig.service.react(BOB, BOB_QUESTION, Reaction.LIKE, False, rig.deadline())

    assert toggle_presses(rig.world.pressed) == []


def test_react_reports_an_unverified_change() -> None:
    rig = writing_rig(standard_world(WorldOptions(reaction_flips=False)))

    record = rig.service.react(BOB, BOB_QUESTION, Reaction.LIKE, False, rig.deadline()).to_record()

    assert record["ok"] is False
    assert record["verified"] is False
    assert record["state"] == "off"
    assert len(toggle_presses(rig.world.pressed)) == 1


def test_react_respects_the_allowlist() -> None:
    rig = writing_rig(standard_world(), allowlist=f"teams:{SELF_THREAD}")

    with pytest.raises(RecipientNotAllowedError):
        rig.service.react(BOB, BOB_QUESTION, Reaction.LIKE, False, rig.deadline())

    assert BOB not in rig.world.pressed


def test_delete_removes_an_own_message() -> None:
    rig = writing_rig(standard_world())

    record = rig.service.delete_message(BOB, MY_ANSWER, rig.deadline()).to_record()

    assert record["action"] == "delete_message"
    assert record["succeeded"] == 1
    assert record["results"][0]["verified"] is True
    assert record["results"][0]["label"] == "Sure, looking now"
    assert rig.world.chat(BOB).messages[-1].deleted is True
    assert ("item", MY_ANSWER, "Delete this message") in rig.world.pressed


def test_delete_also_finds_a_plain_delete_item() -> None:
    world = standard_world(
        WorldOptions(menu_items={MY_ANSWER: ("Reply with quote", "Edit", "Delete")})
    )
    rig = writing_rig(world)

    record = rig.service.delete_message(BOB, MY_ANSWER, rig.deadline()).to_record()

    assert record["succeeded"] == 1
    assert ("item", MY_ANSWER, "Delete") in rig.world.pressed


def test_delete_confirms_the_prompt_when_teams_asks() -> None:
    rig = writing_rig(standard_world(WorldOptions(delete_confirm=True)))

    record = rig.service.delete_message(BOB, MY_ANSWER, rig.deadline()).to_record()

    assert record["succeeded"] == 1
    assert CONFIRM_HANDLE in rig.world.pressed
    assert all(item.id != MY_ANSWER for item in rig.world.chat(BOB).messages)


def test_delete_refuses_messages_from_others() -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(ForbiddenActionError, match="Bob Sample"):
        rig.service.delete_message(BOB, BOB_QUESTION, rig.deadline())

    assert rig.world.menus_opened == []
    assert rig.world.chat(BOB).messages[0].deleted is False


def test_delete_reports_when_the_message_stays() -> None:
    rig = writing_rig(standard_world(WorldOptions(delete_works=False)))

    record = rig.service.delete_message(BOB, MY_ANSWER, rig.deadline()).to_record()

    assert record["failed"] == 1
    assert record["results"][0]["verified"] is False


def test_mark_unread_uses_the_newest_message_from_someone_else() -> None:
    rig = writing_rig(standard_world())

    record = rig.service.mark_unread("Bob Sample", rig.deadline()).to_record()

    assert record["action"] == "mark_unread"
    assert record["succeeded"] == 1
    assert rig.world.menus_opened == [BOB_QUESTION]
    assert rig.world.chat(BOB).unread is True


def test_mark_unread_in_a_self_only_chat_raises_cleanly() -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(ActionNotVerifiedError, match="only on messages from others"):
        rig.service.mark_unread("Alex Example (You)", rig.deadline())

    assert rig.world.menus_opened == []
    assert rig.world.open_thread == ALICE
    assert all(item.id == SELF_NOTE for item in rig.world.chat(SELF_THREAD).messages)


def test_mark_unread_without_the_menu_item_changes_nothing() -> None:
    world = standard_world(WorldOptions(menu_items={BOB_QUESTION: ("Reply with quote",)}))
    rig = writing_rig(world)

    with pytest.raises(ActionNotVerifiedError, match="did not offer"):
        rig.service.mark_unread(BOB, rig.deadline())

    assert world.chat(BOB).unread is False


def test_mark_unread_reports_when_the_row_stays_read() -> None:
    rig = writing_rig(standard_world(WorldOptions(unread_works=False)))

    record = rig.service.mark_unread(BOB, rig.deadline()).to_record()

    assert record["failed"] == 1
    assert record["results"][0]["verified"] is False
