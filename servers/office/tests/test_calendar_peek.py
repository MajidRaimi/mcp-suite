from __future__ import annotations

from datetime import date

import pytest
from fakes.p4_calendar import CalendarDesk, DeskEvent

from office_mcp.features.calendar.models import MyResponse, Responses
from office_mcp.features.calendar.peek import find_peek, my_response, parse_peek, series_prompt
from office_mcp.shared.node import Node

MONDAY = date(2026, 11, 16)


def peek_of(event: DeskEvent) -> Node:
    desk = CalendarDesk([event])
    root = desk.peek(event)
    found = find_peek(Node(role="AXWindow", children=(root,)))
    assert found is not None
    return found


def test_peek_parse_reads_every_field() -> None:
    event = DeskEvent("Project Atlas budget", MONDAY, recurring=True)

    peek = parse_peek(peek_of(event))

    assert peek.subject == "Project Atlas budget"
    assert peek.time_text == "Mon 11/16/2026 10:00 - 10:30"
    assert peek.recurrence == "Series"
    assert peek.location == "Microsoft Teams Meeting"
    assert peek.organizer == "Jane Doe"
    assert peek.invited_text == "Jane Doe invited you."
    assert peek.responses == Responses(accepted=2, not_responded=12)
    assert peek.notes == "Agenda: review the numbers."
    assert peek.has_join_button is True
    assert peek.my_response is MyResponse.NOT_RESPONDED


@pytest.mark.parametrize("state", [MyResponse.ACCEPTED, MyResponse.DECLINED, MyResponse.TENTATIVE])
def test_my_response_from_the_respond_form(state: MyResponse) -> None:
    event = DeskEvent("Weekly sync", MONDAY, response=state)

    assert my_response(peek_of(event)) is state


def test_own_event_without_a_form_is_the_organizer() -> None:
    event = DeskEvent("Focus time", MONDAY, invited=False)

    assert parse_peek(peek_of(event)).my_response is MyResponse.ORGANIZER


def test_find_peek_ignores_trees_without_a_card() -> None:
    assert find_peek(Node(role="AXWindow", children=(Node(role="AXButton", title="Join"),))) is None


def test_series_prompt_is_recognised_only_in_menus() -> None:
    prompt = Node(
        role="AXMenu",
        children=(Node(role="AXMenuItem", title="All events in the series"),),
    )
    unrelated = Node(role="AXMenu", children=(Node(role="AXMenuItem", title="Tentative"),))

    assert series_prompt(Node(role="AXWindow", children=(prompt,))) is prompt
    assert series_prompt(Node(role="AXWindow", children=(unrelated,))) is None
