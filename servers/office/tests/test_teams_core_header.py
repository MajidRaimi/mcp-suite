from __future__ import annotations

import pytest
from fakes.trees import ACTIVITY_APP_ID, chat_header, rail, teams_window

from office_mcp.features.teams.header import (
    ThreadKind,
    open_thread_id,
    participant_count,
    thread_kind,
)
from office_mcp.features.teams.rail import TeamsApp, rail_button, selected_app
from office_mcp.shared.node import Node


def test_open_thread_id_reads_the_chat_header_dom_id() -> None:
    root = teams_window("Chat | Bob Sample", [], [], thread="19:abc@unq.gbl.spaces")

    assert open_thread_id(root) == "19:abc@unq.gbl.spaces"
    assert open_thread_id(teams_window("Chat", [], [])) is None
    assert open_thread_id(Node(role="AXGroup", dom_id="chat-header-")) is None


def test_participant_count_reads_the_roster_and_its_overflow() -> None:
    header = chat_header("19:g@thread.v2", ["Alice Example", "Bob Sample"])
    overflow = Node(role="AXButton", title="+4")
    roster = Node(
        role="AXList", title="Chat participants", children=(*header.children[0].children, overflow)
    )

    assert participant_count(header) == 2
    assert participant_count(Node(role="AXGroup", children=(roster,))) == 6


def test_participant_count_falls_back_to_person_ids_then_none() -> None:
    people = tuple(
        Node(role="AXGroup", dom_id=f"chat-topic-person-8:orgid:{name}") for name in ("a", "b", "a")
    )

    assert participant_count(Node(role="AXGroup", children=people)) == 2
    assert participant_count(Node(role="AXGroup")) is None


def test_group_chat_without_a_roster_has_no_member_count() -> None:
    people = tuple(
        Node(role="AXGroup", dom_id=f"chat-topic-person-8:orgid:{name}") for name in "abcd"
    )
    header = Node(role="AXGroup", dom_id="chat-header-19:allhands@thread.v2", children=people)
    direct = Node(role="AXGroup", dom_id="chat-header-19:a_b@unq.gbl.spaces", children=people)

    assert participant_count(header) is None
    assert participant_count(direct) == 4


@pytest.mark.parametrize(
    ("thread_id", "kind"),
    [
        ("48:notes", ThreadKind.NOTES),
        ("19:general@thread.tacv2", ThreadKind.CHANNEL),
        ("19:old@thread.skype", ThreadKind.CHANNEL),
        ("19:meeting_abc@thread.v2", ThreadKind.MEETING),
        ("19:group@thread.v2", ThreadKind.GROUP),
        ("19:abc_def@unq.gbl.spaces", ThreadKind.CHAT),
    ],
)
def test_thread_kind_table(thread_id: str, kind: ThreadKind) -> None:
    assert thread_kind(thread_id) is kind


def test_selected_app_reads_the_rail_toggle_value() -> None:
    assert selected_app(rail()) is TeamsApp.CHAT
    assert selected_app(rail(ACTIVITY_APP_ID)) is TeamsApp.ACTIVITY
    assert selected_app(rail(None)) is None
    button = rail_button(rail(), TeamsApp.ACTIVITY)
    assert button is not None
    assert button.handle == "rail:Activity"
    assert rail_button(rail(), TeamsApp.CALENDAR) is None
