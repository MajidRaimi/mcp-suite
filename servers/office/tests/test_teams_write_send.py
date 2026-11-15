from __future__ import annotations

import pytest
from fakes.guard import RecordingGuard
from fakes.p5_rig import (
    ALICE,
    BIG_GROUP,
    BOB,
    BOB_QUESTION,
    CHANNEL,
    GROUP,
    MEETING,
    OWNER_APP,
    SELF_THREAD,
    WritingRig,
    standard_chats,
    standard_world,
    writing_rig,
)
from fakes.p5_world import PASTE, SEND_KEYS, TeamsWorld, WorldOptions

from office_mcp.features.teams.writing.send import (
    DISCARDED_NOTE,
    KEPT_NOTE,
    STILL_IN_BOX,
    SWITCHED_NOTE,
    SendRequest,
)
from office_mcp.shared.drafts import Channel, Draft, Review, Severity, Violation
from office_mcp.shared.errors import (
    ComposeBusyError,
    DuplicateSendError,
    ForbiddenActionError,
    PersonaViolationError,
    RecipientNotAllowedError,
    SendRefusedError,
    TargetError,
    TextEntryError,
    UiTimeoutError,
)
from office_mcp.shared.events import Shortcut
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.results import UNCONFIRMED_NOTE, SendStatus

HELLO = "Thanks, I will check it today."


def request(
    chat: str,
    text: str = HELLO,
    group_ok: bool = False,
    allow_language_mismatch: bool = False,
    allow_duplicate: bool = False,
    dry_run: bool = False,
) -> SendRequest:
    return SendRequest(
        chat=chat,
        text=text,
        group_ok=group_ok,
        allow_language_mismatch=allow_language_mismatch,
        allow_duplicate=allow_duplicate,
        dry_run=dry_run,
    )


def pasted(rig: WritingRig) -> list[str]:
    return list(rig.world.pasteboard.writes)


def posted(world: TeamsWorld, thread: str) -> list[str]:
    return [item.text for item in world.chat(thread).messages if item.mine]


def violation_review(draft: Draft) -> Review:
    broken = Violation("characters", "forbidden character", Severity.ERROR)
    return Review(draft=draft, violations=(broken,), fixes=(), persona_version="test")


class StrictGuard(RecordingGuard):
    def verify_written(self, intended: Draft, body_read: str, subject_read: str) -> None:
        super().verify_written(intended, body_read, subject_read)
        raise PersonaViolationError(violation_review(intended).violations)


def test_send_types_once_commits_once_and_verifies_by_id_and_text() -> None:
    rig = writing_rig(standard_world())

    result = rig.service.send(request("Bob Sample"), rig.deadline())

    record = result.to_record()
    new_message = rig.world.chat(BOB).messages[-1]
    assert record["status"] == "sent"
    assert record["verified"] is True
    assert record["message_ref"] == new_message.id
    assert int(new_message.id) > int(BOB_QUESTION)
    assert record["targets"] == [BOB]
    assert record["chat"] == "Bob Sample"
    assert record["text_sent"] == HELLO
    assert rig.world.commits == 1
    assert pasted(rig) == [HELLO]
    assert rig.events.shortcuts() == [PASTE]
    assert rig.world.box_text(BOB) == ""
    assert rig.world.open_thread == ALICE
    assert rig.workspace.frontmost == OWNER_APP
    assert result.focus is FocusOutcome.RESTORED
    assert [entry[1] for entry in rig.guard.verified] == [HELLO]


def test_send_reviews_with_the_newest_message_from_someone_else_as_context() -> None:
    rig = writing_rig(standard_world())

    rig.service.send(request(BOB, allow_language_mismatch=True), rig.deadline())

    first, second = rig.guard.reviewed
    assert first.context == ""
    assert second.context == "Can you review the Project Atlas budget?"
    assert second.context_author == "Bob Sample"
    assert second.recipients == (BOB,)
    assert second.allow_language_mismatch is True


def test_multi_line_text_is_pasted_and_the_clipboard_is_restored() -> None:
    rig = writing_rig(standard_world())
    text = "Line one\nLine two"

    result = rig.service.send(request(BOB, text), rig.deadline())

    assert result.status is SendStatus.SENT
    assert rig.world.pasteboard.writes == [text]
    assert rig.events.shortcuts() == [PASTE]
    assert rig.world.pasteboard.text() == "owner clipboard"
    assert rig.world.commits == 1
    assert posted(rig.world, BOB)[-1] == text


def test_early_send_during_paste_aborts_without_a_commit() -> None:
    rig = writing_rig(standard_world(WorldOptions(send_on_newline=True)))

    with pytest.raises(SendRefusedError, match="sent part of the text early") as caught:
        rig.service.send(request(BOB, "First line\nSecond line"), rig.deadline())

    assert rig.world.commits == 0
    assert SEND_KEYS not in rig.events.shortcuts()
    assert "rest of the text" in "\n".join(caught.value.__notes__)
    assert rig.world.open_thread == ALICE


@pytest.mark.parametrize("query", ["Bob", "bob sam", "Sample", "19:bob"])
def test_partial_names_are_refused_before_anything_is_opened(query: str) -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(TargetError, match="exact chat name"):
        rig.service.send(request(query), rig.deadline())

    assert rig.world.pressed == []
    assert rig.world.commits == 0


@pytest.mark.parametrize("query", [CHANNEL, "19:old@thread.skype"])
def test_channel_threads_are_refused(query: str) -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(ForbiddenActionError, match="channels"):
        rig.service.send(request(query, dry_run=True), rig.deadline())

    assert rig.world.pressed == []


def test_meeting_chats_need_group_ok() -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(ForbiddenActionError, match="meeting chat"):
        rig.service.send(request("Weekly sync"), rig.deadline())

    assert MEETING not in rig.world.pressed
    result = rig.service.send(request("Weekly sync", group_ok=True), rig.deadline())
    assert result.status is SendStatus.SENT
    assert rig.world.commits == 1


def test_large_groups_need_group_ok_and_small_groups_do_not() -> None:
    rig = writing_rig(standard_world(), max_members=25)

    with pytest.raises(ForbiddenActionError, match="30 members"):
        rig.service.send(request(BIG_GROUP), rig.deadline())

    assert pasted(rig) == []
    assert rig.service.send(request(GROUP), rig.deadline()).status is SendStatus.SENT
    forced = rig.service.send(request(BIG_GROUP, group_ok=True), rig.deadline())
    assert forced.status is SendStatus.SENT
    assert forced.to_record()["members"] == 30
    assert rig.world.commits == 2


def test_persona_rejection_happens_before_any_ui() -> None:
    guard = RecordingGuard([violation_review(Draft(channel=Channel.TEAMS, body=HELLO))])
    rig = writing_rig(standard_world(), guard=guard)

    with pytest.raises(PersonaViolationError):
        rig.service.send(request(BOB), rig.deadline())

    assert rig.world.snapshots == 0
    assert pasted(rig) == []


def test_language_rejection_after_opening_leaves_the_box_untouched() -> None:
    rig = writing_rig(standard_world())
    clean = Review(draft=Draft(Channel.TEAMS, HELLO), violations=(), fixes=(), persona_version="t")
    rig.guard.scripted = [clean, violation_review(Draft(Channel.TEAMS, HELLO))]

    with pytest.raises(PersonaViolationError):
        rig.service.send(request(BOB), rig.deadline())

    assert BOB in rig.world.pressed
    assert pasted(rig) == []
    assert rig.world.commits == 0
    assert rig.world.open_thread == ALICE


def test_allowlist_blocks_other_chats_before_opening_them() -> None:
    rig = writing_rig(standard_world(), allowlist=f"teams:{SELF_THREAD}")

    with pytest.raises(RecipientNotAllowedError):
        rig.service.send(request(BOB, dry_run=True), rig.deadline())

    assert BOB not in rig.world.pressed
    assert rig.service.send(request(SELF_THREAD), rig.deadline()).status is SendStatus.SENT


def test_teams_self_entry_allows_only_the_configured_thread() -> None:
    rig = writing_rig(standard_world(), allowlist="teams:self")

    assert rig.service.send(request("Alex Example (You)"), rig.deadline()).verified
    with pytest.raises(RecipientNotAllowedError):
        rig.service.send(request("Alice Example"), rig.deadline())


def test_a_draft_in_the_box_blocks_the_send_and_is_kept() -> None:
    rig = writing_rig(standard_world(bob_draft="half written note"))

    with pytest.raises(ComposeBusyError):
        rig.service.send(request(BOB), rig.deadline())

    assert rig.world.box_text(BOB) == "half written note"
    assert pasted(rig) == []
    assert rig.world.commits == 0
    assert rig.workspace.activations == []


def test_dry_run_checks_everything_and_types_nothing() -> None:
    rig = writing_rig(standard_world())

    result = rig.service.send(request(BOB, dry_run=True), rig.deadline())

    record = result.to_record()
    assert record["status"] == "dry_run"
    assert record["ok"] is True
    assert record["text_sent"] == HELLO
    assert pasted(rig) == []
    assert rig.world.commits == 0
    assert rig.workspace.activations == []
    assert rig.world.open_thread == ALICE


def test_the_same_text_as_a_recent_own_message_is_a_duplicate() -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(DuplicateSendError, match="1790000002000"):
        rig.service.send(request(BOB, "sure,  LOOKING now"), rig.deadline())

    assert rig.world.commits == 0
    forced = rig.service.send(
        request(BOB, "Sure, looking now", allow_duplicate=True), rig.deadline()
    )
    assert forced.status is SendStatus.SENT


def test_the_send_guard_blocks_a_quick_resend_even_when_the_chat_moved_on() -> None:
    rig = writing_rig(standard_world(WorldOptions(posted_text="something else")))

    first = rig.service.send(request(BOB), rig.deadline())

    assert first.status is SendStatus.UNCONFIRMED
    with pytest.raises(DuplicateSendError, match="s ago"):
        rig.service.send(request(BOB), rig.deadline())
    assert rig.world.commits == 1


def test_a_substituted_character_is_caught_and_the_box_is_left_alone() -> None:
    rig = writing_rig(standard_world(WorldOptions(substitutions={"-": "\u2014"})))

    with pytest.raises(TextEntryError) as caught:
        rig.service.send(request(BOB, "plan - then go"), rig.deadline())

    assert rig.world.commits == 0
    assert KEPT_NOTE in caught.value.__notes__
    assert rig.world.box_text(BOB) == "plan \u2014 then go"


def test_a_persona_hit_in_the_read_back_text_clears_exactly_what_was_typed() -> None:
    rig = writing_rig(standard_world(), guard=StrictGuard())

    with pytest.raises(PersonaViolationError) as caught:
        rig.service.send(request(BOB), rig.deadline())

    assert rig.world.commits == 0
    assert DISCARDED_NOTE in caught.value.__notes__
    assert rig.world.box_text(BOB) == ""


def test_a_chat_switch_before_send_refuses_and_keeps_the_text_without_any_key() -> None:
    world = standard_world()
    rig = writing_rig(world)

    def paste_then_switch(pid: int, shortcut: Shortcut) -> None:
        world.on_event(pid, shortcut)
        if shortcut == PASTE:
            world.open_thread = ALICE

    rig.events.on_press = paste_then_switch

    with pytest.raises(TargetError, match="does not show the chat") as caught:
        rig.service.send(request(BOB), rig.deadline())

    assert world.commits == 0
    assert KEPT_NOTE in caught.value.__notes__
    assert rig.events.shortcuts() == [PASTE]
    assert world.box_text(ALICE) == ""


def test_a_chat_switch_during_the_clear_presses_no_backspace() -> None:
    world = standard_world()
    rig = writing_rig(world, guard=StrictGuard())
    alice_draft = "Owner draft for Alice"
    world.set_box(world.chat(ALICE), alice_draft)
    bob_box = world.chat(BOB).box

    def select_then_switch(value: object) -> None:
        bob_box.attributes["AXSelectedTextRange"] = value
        world.open_thread = ALICE

    bob_box.on_set["AXSelectedTextRange"] = select_then_switch

    with pytest.raises(PersonaViolationError) as caught:
        rig.service.send(request(BOB), rig.deadline())

    assert world.commits == 0
    assert SWITCHED_NOTE in caught.value.__notes__
    assert rig.events.shortcuts() == [PASTE]
    assert world.box_text(ALICE) == alice_draft
    assert world.box_text(BOB) == HELLO


def test_running_out_of_time_before_send_clears_the_box() -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(UiTimeoutError, match="nothing was sent") as caught:
        rig.service.send(request(BOB), rig.deadline(0.0))

    assert rig.world.commits == 0
    assert DISCARDED_NOTE in caught.value.__notes__


def test_a_send_that_never_shows_is_unconfirmed_not_an_error() -> None:
    rig = writing_rig(standard_world(WorldOptions(send_works=False)))

    result = rig.service.send(request(BOB), rig.deadline())

    record = result.to_record()
    assert record["status"] == "unconfirmed"
    assert record["ok"] is False
    assert record["verified"] is False
    assert UNCONFIRMED_NOTE in record["notes"]
    assert STILL_IN_BOX in record["notes"]
    assert rig.world.commits == 1


def test_a_new_message_with_other_text_is_unconfirmed() -> None:
    rig = writing_rig(standard_world(WorldOptions(posted_text="garbled")))

    record = rig.service.send(request(BOB), rig.deadline()).to_record()

    assert record["status"] == "unconfirmed"
    assert any("text differs" in note for note in record["notes"])
    assert rig.world.commits == 1


def test_without_a_send_button_the_shortcut_is_the_single_commit() -> None:
    rig = writing_rig(standard_world(WorldOptions(send_button=False)))

    result = rig.service.send(request(BOB), rig.deadline())

    assert result.status is SendStatus.SENT
    assert rig.world.keys == [SEND_KEYS]
    assert rig.world.commits == 1


def test_a_box_that_never_takes_keyboard_focus_is_refused_before_pasting() -> None:
    world = standard_world(WorldOptions(send_button=False))
    rig = writing_rig(world)
    world.chat(BOB).box.on_set["AXFocused"] = lambda value: None

    with pytest.raises(TextEntryError, match="nothing was pasted"):
        rig.service.send(request(BOB), rig.deadline())

    assert rig.world.keys == []
    assert rig.world.commits == 0
    assert rig.world.box_text(BOB) == ""


def test_an_empty_or_long_text_is_refused_before_any_ui() -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(TextEntryError, match="empty"):
        rig.service.send(request(BOB, "   "), rig.deadline())
    with pytest.raises(TextEntryError, match="4000"):
        rig.service.send(request(BOB, "x" * 4001), rig.deadline())

    assert rig.world.snapshots == 0


def unlisted_self_world(open_thread: str) -> TeamsWorld:
    chats = standard_chats("")
    for chat in chats:
        chat.listed = chat.thread != SELF_THREAD
    return TeamsWorld(chats, open_thread=open_thread)


def test_an_open_chat_missing_from_the_list_is_targeted_by_its_thread_id() -> None:
    rig = writing_rig(unlisted_self_world(SELF_THREAD), allowlist=f"teams:{SELF_THREAD}")

    result = rig.service.send(request(SELF_THREAD), rig.deadline())

    assert result.status is SendStatus.SENT
    assert posted(rig.world, SELF_THREAD)[-1] == HELLO
    assert rig.world.commits == 1
    assert rig.world.open_thread == SELF_THREAD


def test_an_unlisted_chat_that_is_not_open_is_refused_before_any_ui() -> None:
    rig = writing_rig(unlisted_self_world(ALICE), allowlist=f"teams:{SELF_THREAD}")

    with pytest.raises(TargetError):
        rig.service.send(request(SELF_THREAD), rig.deadline())

    assert rig.world.commits == 0
    assert rig.world.open_thread == ALICE


def test_an_open_unlisted_chat_is_not_matched_by_its_name() -> None:
    rig = writing_rig(unlisted_self_world(SELF_THREAD), allowlist="teams:self")

    with pytest.raises(TargetError):
        rig.service.send(request("Alex Example (You)"), rig.deadline())

    assert rig.world.commits == 0
