from __future__ import annotations

import pytest
from fakes.guard import RecordingGuard
from fakes.p5_rig import ALICE, BOB, BOB_QUESTION, MY_ANSWER, standard_world, writing_rig
from fakes.p5_world import WorldOptions, other

from office_mcp.features.teams.writing.send import DISCARDED_NOTE, SendRequest
from office_mcp.shared.drafts import Draft
from office_mcp.shared.errors import (
    ComposeError,
    PersonaViolationError,
    TargetError,
    TargetNotFoundError,
)
from office_mcp.shared.results import SendStatus

ANSWER = "Approved, thanks."


def reply(message_id: str, text: str = ANSWER) -> SendRequest:
    return SendRequest(chat="Bob Sample", text=text, reply_to=message_id)


class RejectReadBack(RecordingGuard):
    def verify_written(self, intended: Draft, body_read: str, subject_read: str) -> None:
        super().verify_written(intended, body_read, subject_read)
        raise PersonaViolationError(())


def test_reply_quotes_the_message_and_verifies_the_new_one() -> None:
    rig = writing_rig(standard_world())

    result = rig.service.send(reply(BOB_QUESTION), rig.deadline())

    record = result.to_record()
    newest = rig.world.chat(BOB).messages[-1]
    assert record["status"] == "sent"
    assert record["message_ref"] == newest.id
    assert record["quoted"] == {"message_id": BOB_QUESTION, "author": "Bob Sample"}
    assert newest.text.startswith("Begin quote, Bob Sample")
    assert newest.text.endswith(ANSWER)
    assert rig.world.menus_opened == [BOB_QUESTION]
    assert rig.world.commits == 1
    assert [entry[1] for entry in rig.guard.verified] == [ANSWER]
    assert rig.world.open_thread == ALICE


def test_reply_works_when_the_quote_sits_outside_the_box_value() -> None:
    rig = writing_rig(standard_world(WorldOptions(quote_in_value=False)))

    result = rig.service.send(reply(MY_ANSWER), rig.deadline())

    assert result.status is SendStatus.SENT
    assert result.to_record()["quoted"]["author"] == "Alex Example"


def test_a_missing_quote_stops_before_typing() -> None:
    rig = writing_rig(standard_world(WorldOptions(quote_works=False)))

    with pytest.raises(ComposeError, match="did not add a quote"):
        rig.service.send(reply(BOB_QUESTION), rig.deadline())

    assert rig.world.commits == 0
    assert rig.world.box_text(BOB) == ""


def test_a_failure_after_quoting_clears_the_quote_and_the_text() -> None:
    rig = writing_rig(standard_world(), guard=RejectReadBack())

    with pytest.raises(PersonaViolationError) as caught:
        rig.service.send(reply(BOB_QUESTION), rig.deadline())

    assert DISCARDED_NOTE in caught.value.__notes__
    assert rig.world.box_text(BOB) == ""
    assert rig.world.commits == 0


def test_reply_scrolls_up_to_find_an_older_message() -> None:
    world = standard_world()
    world.chat(BOB).older = [other(1789999999000, "Bob Sample", "Older budget question")]
    rig = writing_rig(world)

    result = rig.service.send(reply("1789999999000"), rig.deadline())

    assert result.status is SendStatus.SENT
    assert world.scrolled
    assert result.to_record()["quoted"]["message_id"] == "1789999999000"


def test_an_unknown_message_is_not_found_and_nothing_is_typed() -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(TargetNotFoundError, match="1700000000000"):
        rig.service.send(reply("1700000000000"), rig.deadline())

    assert rig.world.commits == 0
    assert rig.world.menus_opened == []


@pytest.mark.parametrize("message_id", ["latest", "12ab", ""])
def test_a_malformed_message_id_is_refused_before_any_ui(message_id: str) -> None:
    rig = writing_rig(standard_world())

    with pytest.raises(TargetError, match="message id from teams_read"):
        rig.service.send(reply(message_id), rig.deadline())

    assert rig.world.snapshots == 0
