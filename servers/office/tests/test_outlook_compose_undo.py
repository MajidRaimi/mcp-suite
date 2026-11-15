from __future__ import annotations

import pytest
from fakes.p3_compose import FakeCompose, compose_rig

from office_mcp.features.outlook.compose.service import NewMail
from office_mcp.shared.errors import ActionNotVerifiedError, TargetNotFoundError

SUBJECT = "[mcp-suite test] T2"


def test_undo_within_the_window_discards_the_reopened_draft() -> None:
    rig = compose_rig(compose=FakeCompose(undo_enabled=True, reopen_subject=SUBJECT))
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)
    rig.service.send(NewMail(to=("jane@example.com",), subject=SUBJECT, body="Hi"), rig.deadline())
    rig.clock.advance(5)

    result = rig.service.undo_send().to_record()

    assert result["ok"] is True
    assert result["undone_subject"] == SUBJECT
    assert "Deleted Items" in result["detail"]
    assert rig.compose.calls[-2:] == ["undo_send", "discard"]
    assert rig.compose.open == []
    assert rig.sends.last(60) is None


def test_undo_after_thirty_seconds_does_nothing() -> None:
    rig = compose_rig(compose=FakeCompose(undo_enabled=True, reopen_subject=SUBJECT))
    rig.service.send(NewMail(to=("jane@example.com",), subject=SUBJECT, body="Hi"), rig.deadline())
    rig.clock.advance(31)
    calls = list(rig.compose.calls)
    scopes = list(rig.scope_log)

    with pytest.raises(TargetNotFoundError, match="nothing to undo"):
        rig.service.undo_send()

    assert rig.compose.calls == calls
    assert rig.scope_log == scopes


def test_undo_without_any_send_does_nothing() -> None:
    rig = compose_rig()

    with pytest.raises(TargetNotFoundError):
        rig.service.undo_send()

    assert rig.compose.calls == []


def test_undo_refused_when_outlook_no_longer_offers_it() -> None:
    rig = compose_rig(compose=FakeCompose(undo_enabled=False))
    rig.service.send(NewMail(to=("jane@example.com",), subject=SUBJECT, body="Hi"), rig.deadline())

    with pytest.raises(ActionNotVerifiedError, match="no longer offers"):
        rig.service.undo_send()

    assert "undo_send" not in rig.compose.calls


def test_undo_without_a_reopened_compose_is_not_verified() -> None:
    rig = compose_rig(compose=FakeCompose(undo_enabled=True, reopen_subject=None))
    rig.service.send(NewMail(to=("jane@example.com",), subject=SUBJECT, body="Hi"), rig.deadline())

    with pytest.raises(ActionNotVerifiedError, match="no compose window reopened"):
        rig.service.undo_send()


def test_undo_leaves_a_compose_with_another_subject_open() -> None:
    rig = compose_rig(compose=FakeCompose(undo_enabled=True, reopen_subject="Something else"))
    rig.service.send(NewMail(to=("jane@example.com",), subject=SUBJECT, body="Hi"), rig.deadline())

    with pytest.raises(ActionNotVerifiedError, match="left open"):
        rig.service.undo_send()

    assert rig.compose.calls[-1] == "undo_send"
    assert len(rig.compose.open) == 1


def test_undo_is_refused_when_another_mail_went_out_after_ours() -> None:
    rig = compose_rig(compose=FakeCompose(undo_enabled=True, reopen_subject="Owner mail"))
    rig.compose.on_send = lambda: rig.deliver_to_sent(SUBJECT)
    rig.service.send(NewMail(to=("jane@example.com",), subject=SUBJECT, body="Hi"), rig.deadline())
    rig.deliver_to_sent("Owner mail", time_label="9:30 AM")
    rig.clock.advance(5)

    with pytest.raises(ActionNotVerifiedError, match="nothing was undone"):
        rig.service.undo_send()

    assert "undo_send" not in rig.compose.calls
    assert rig.compose.open == []
