from __future__ import annotations

import pytest
from fakes.clock import fake_waiter
from fakes.p3_compose import FakeCompose

from office_mcp.features.outlook.compose.gateway import ComposeSurface, RecipientField
from office_mcp.features.outlook.compose.session import AFTER_BODY, FAIL_POINTS, ComposeSession
from office_mcp.shared.errors import ComposeError, UiTimeoutError, error_text


def fill(session: ComposeSession) -> None:
    session.require_empty_body()
    session.fill(
        ((RecipientField.TO, ("jane@example.com",)),),
        "[mcp-suite test] abort",
        "This must never be sent.",
    )
    session.commit()


@pytest.mark.parametrize("point", FAIL_POINTS)
def test_fail_at_discards_without_pressing_send(point: str) -> None:
    compose = FakeCompose()

    with (
        pytest.raises(ComposeError, match=point) as caught,
        ComposeSession(compose, compose.open_new, fake_waiter(), fail_at=point) as session,
    ):
        fill(session)

    assert compose.send_presses == 0
    assert compose.discards == 1
    assert compose.open == []
    assert "a copy is in Deleted Items" in error_text(caught.value)


def test_unknown_fail_point_is_refused() -> None:
    compose = FakeCompose()

    with pytest.raises(ComposeError, match="fail_at"):
        ComposeSession(compose, compose.open_new, fake_waiter(), fail_at="halfway")


def test_commit_presses_send_only_once() -> None:
    compose = FakeCompose(close_on_send=False)

    with (
        pytest.raises(ComposeError, match="already pressed"),
        ComposeSession(compose, compose.open_new, fake_waiter()) as session,
    ):
        session.commit()
        session.commit()

    assert compose.send_presses == 1
    assert compose.discards == 0


def test_open_failure_discards_a_late_compose_window() -> None:
    compose = FakeCompose(open_error=UiTimeoutError("no window in time"), stray_on_error=True)

    with (
        pytest.raises(UiTimeoutError) as caught,
        ComposeSession(compose, compose.open_new, fake_waiter()),
    ):
        pytest.fail("the body must not run")

    assert compose.discards == 1
    assert compose.open == []
    assert "a copy is in Deleted Items" in error_text(caught.value)


def test_failed_discard_is_reported() -> None:
    compose = FakeCompose(
        discard_error=UiTimeoutError("the unsent draft did not close after Discard")
    )

    with (
        pytest.raises(ComposeError) as caught,
        ComposeSession(compose, compose.open_new, fake_waiter(), fail_at=AFTER_BODY) as s,
    ):
        fill(s)

    text = error_text(caught.value)
    assert "could not be discarded" in text
    assert "close it in Outlook without sending" in text
    assert compose.send_presses == 0


def test_observe_requires_a_commit() -> None:
    compose = FakeCompose()

    with (
        pytest.raises(ComposeError, match="not pressed"),
        ComposeSession(compose, compose.open_new, fake_waiter()) as session,
    ):
        session.observe()

    assert compose.discards == 1


def test_open_failure_leaves_a_late_compose_that_holds_text() -> None:
    compose = FakeCompose(
        open_error=UiTimeoutError("no window in time"),
        stray_on_error=True,
        stray_subject="Owner draft",
    )

    with (
        pytest.raises(UiTimeoutError) as caught,
        ComposeSession(compose, compose.open_new, fake_waiter()),
    ):
        pytest.fail("the body must not run")

    assert compose.discards == 0
    assert len(compose.open) == 1
    assert "could not be proven" in error_text(caught.value)


def test_open_failure_with_two_late_composes_discards_neither() -> None:
    compose = FakeCompose(open_error=UiTimeoutError("no window in time"), stray_on_error=True)

    def open_two() -> ComposeSurface:
        compose.open.append(ComposeSurface(handle="owner", window="owner", inline=False))
        return compose.open_new()

    with (
        pytest.raises(UiTimeoutError) as caught,
        ComposeSession(compose, open_two, fake_waiter()),
    ):
        pytest.fail("the body must not run")

    assert compose.discards == 0
    assert len(compose.open) == 2
    assert "could not be proven" in error_text(caught.value)
