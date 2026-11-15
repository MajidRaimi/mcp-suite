from __future__ import annotations

from datetime import date

import pytest
from fakes.clock import FakeClock
from fakes.guard import RecordingGuard
from fakes.p4_calendar import CalendarDesk, CalendarRig, DeskEvent, calendar_rig
from fakes.workspace import FakeWorkspace

from office_mcp.features.calendar.models import MyResponse
from office_mcp.features.calendar.rsvp import rsvp_request
from office_mcp.features.teams.rail import TeamsApp
from office_mcp.shared.drafts import Channel, Draft, Review, Severity, Violation
from office_mcp.shared.errors import (
    ComposeBusyError,
    ElementNotFoundError,
    PersonaViolationError,
    RecipientNotAllowedError,
    SendRefusedError,
    TargetError,
    TextEntryError,
)
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.recipients import RecipientGuard, parse_allowlist
from office_mcp.shared.runner import Deadline

MONDAY = date(2026, 11, 16)
OWNER_APP = 100


def deadline() -> Deadline:
    return Deadline.after(FakeClock(), 45.0)


def listed_rig(
    response: MyResponse = MyResponse.NOT_RESPONDED,
    recipients: RecipientGuard | None = None,
    guard: RecordingGuard | None = None,
) -> tuple[CalendarRig, str]:
    desk = CalendarDesk([DeskEvent("Project Atlas budget", MONDAY, response=response)])
    rig = calendar_rig(desk, recipients, guard, workspace=FakeWorkspace(frontmost=OWNER_APP))
    listed = rig.service.events("2026-11-16", 1, False, 10, deadline())
    desk.pressed.clear()
    desk.snapshots = 0
    return rig, listed[0].id


def respond(
    rig: CalendarRig,
    event: str,
    response: str,
    notify: bool = True,
    comment: str | None = None,
    dry_run: bool = False,
) -> dict[str, object]:
    request = rsvp_request(event, response, notify, comment, dry_run)
    return rig.service.respond(request, deadline()).to_record()


def commits(rig: CalendarRig) -> list[str]:
    return [item for item in rig.desk.pressed if item in ("accept", "decline", "item:Tentative")]


def test_accept_presses_once_verifies_and_restores_everything() -> None:
    rig, event = listed_rig()

    record = respond(rig, event, "accept")

    assert record["ok"] is True
    assert record["verified"] is True
    assert record["notified"] is True
    assert record["persona"] is None
    assert record["focus_restored"] == FocusOutcome.RESTORED.value
    assert commits(rig) == ["accept"]
    assert rig.desk.events[0].response is MyResponse.ACCEPTED
    assert rig.desk.app is TeamsApp.CHAT
    assert rig.desk.peek_for is None
    assert rig.workspace.frontmost == OWNER_APP


def test_tentative_goes_through_more_options() -> None:
    rig, event = listed_rig()

    record = respond(rig, event, "tentative")

    assert record["verified"] is True
    assert rig.desk.pressed.index("more") < rig.desk.pressed.index("item:Tentative")
    assert commits(rig) == ["item:Tentative"]
    assert rig.desk.events[0].response is MyResponse.TENTATIVE


def test_notify_false_unticks_email_organizer_first() -> None:
    rig, event = listed_rig()

    record = respond(rig, event, "decline", notify=False)

    assert record["notified"] is False
    assert record["verified"] is True
    assert rig.desk.pressed.index("email") < rig.desk.pressed.index("decline")


def test_same_response_is_a_no_op() -> None:
    rig, event = listed_rig(MyResponse.ACCEPTED)

    record = respond(rig, event, "accept")

    assert record["ok"] is True
    assert "already accepted" in str(record["detail"])
    assert record["notified"] is False
    assert commits(rig) == []


def test_refused_under_allowlist_with_notify_before_any_ui() -> None:
    allowlist = RecipientGuard(parse_allowlist("jane@example.com,teams:self"), "48:notes")
    rig, event = listed_rig(recipients=allowlist)

    with pytest.raises(RecipientNotAllowedError, match="notify=false"):
        respond(rig, event, "accept")

    assert rig.desk.snapshots == 0
    record = respond(rig, event, "accept", notify=False)
    assert record["verified"] is True


def test_persona_gate_on_the_comment_uses_the_rsvp_channel() -> None:
    violation = Violation("characters", "forbidden character", Severity.ERROR)
    blocked = Review(Draft(Channel.RSVP, "x"), (violation,), (), "test")
    guard = RecordingGuard([blocked])
    rig, event = listed_rig(guard=guard)

    with pytest.raises(PersonaViolationError):
        respond(rig, event, "accept", comment="See you there \u2014 thanks")

    assert guard.reviewed[0].channel is Channel.RSVP
    assert guard.reviewed[0].body == "See you there \u2014 thanks"
    assert rig.desk.snapshots == 0


def test_comment_is_written_checked_and_sent() -> None:
    rig, event = listed_rig()

    record = respond(rig, event, "accept", comment="See you there")

    assert rig.text.writes == ["See you there"]
    assert rig.guard.verified[0][1] == "See you there"
    assert record["persona"] == {"version": "test", "fixes": [], "warnings": []}
    assert commits(rig) == ["accept"]


def test_comment_that_reads_back_differently_is_not_sent() -> None:
    rig, event = listed_rig()
    rig.text.typo = "See you there."

    with pytest.raises(TextEntryError):
        respond(rig, event, "accept", comment="See you there")

    assert commits(rig) == []
    assert rig.desk.message == "See you there."
    assert rig.desk.events[0].response is MyResponse.NOT_RESPONDED


def test_failure_after_typing_clears_only_the_typed_comment() -> None:
    rig, event = listed_rig()
    rig.desk.more_items = ("Follow",)

    with pytest.raises(ElementNotFoundError, match="nothing was sent"):
        respond(rig, event, "tentative", comment="See you there")

    assert rig.text.cleared == ["See you there"]
    assert rig.desk.message == ""
    assert commits(rig) == []
    assert rig.desk.more_menu is False


def test_busy_message_box_refuses() -> None:
    rig, event = listed_rig()
    rig.desk.message = "half typed by the owner"

    with pytest.raises(ComposeBusyError):
        respond(rig, event, "accept")

    assert commits(rig) == []
    assert rig.desk.message == "half typed by the owner"


def test_series_prompt_is_closed_and_refused() -> None:
    rig, event = listed_rig()
    rig.desk.series_prompt = True

    with pytest.raises(SendRefusedError, match="nothing was sent"):
        respond(rig, event, "decline")

    assert rig.desk.series_menu is False
    assert rig.desk.events[0].response is MyResponse.NOT_RESPONDED


def test_unconfirmed_response_is_not_ok_and_says_check_first() -> None:
    rig, event = listed_rig()
    rig.desk.apply_response = False

    record = respond(rig, event, "accept")

    assert record["ok"] is False
    assert record["verified"] is False
    assert "before retrying" in str(record["detail"])
    assert commits(rig) == ["accept"]


def test_dry_run_runs_checks_and_touches_nothing() -> None:
    rig, event = listed_rig()

    record = respond(rig, event, "accept", comment="See you there", dry_run=True)

    assert record["ok"] is True
    assert record["verified"] is False
    assert record["notified"] is False
    assert rig.desk.snapshots == 0
    assert rig.guard.reviewed[0].channel is Channel.RSVP


@pytest.mark.parametrize(
    ("response", "notify", "comment", "message"),
    [
        ("maybe", True, None, "accept, tentative or decline"),
        ("accept", False, "see you", "notify=true"),
        ("accept", True, "x" * 2001, "2000"),
    ],
)
def test_request_validation(response: str, notify: bool, comment: str | None, message: str) -> None:
    with pytest.raises(TargetError, match=message):
        rsvp_request("0123456789ab", response, notify, comment, False)
