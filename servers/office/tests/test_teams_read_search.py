from __future__ import annotations

import pytest
from conftest import TODAY
from fakes.drivers import fake_runtime, teams_rig
from fakes.p6_trees import ScriptedTeamsGateway, search_box, search_popup, teams_screen

from office_mcp.features.teams.rail import TeamsApp
from office_mcp.features.teams.reading.factory import build_teams_reading
from office_mcp.features.teams.reading.service import TeamsReadingService
from office_mcp.features.teams.search.models import ResultType, SearchResult
from office_mcp.features.teams.search.parse import parse_results
from office_mcp.shared.errors import TargetError
from office_mcp.shared.events import Key, Shortcut
from office_mcp.shared.node import Node

TYPEAHEAD = (
    "Messages",
    "Files",
    "Group chat Project Atlas, Jane Doe, Omar Haddad",
    "Suggestion atlas budget",
    "Powerpoint file result Atlas plan.pptx, Created by Jane Doe",
    "Recent messages",
    "Message from Omar Haddad at Omar Haddad: the atlas numbers are in",
    "Omar Haddad",
    "Yesterday",
    "Person Alex Example",
)


def idle() -> Node:
    return teams_screen(TeamsApp.CHAT, search_box())


def typed(query: str, *texts: str) -> Node:
    return teams_screen(TeamsApp.CHAT, search_box(query), search_popup(*texts))


def scripted(*texts: str) -> ScriptedTeamsGateway:
    gateway = ScriptedTeamsGateway([idle()])

    def on_value(_node: Node, value: str) -> None:
        gateway.show(typed(value, *texts) if value else idle())

    gateway.on_value = on_value
    return gateway


def reading(gateway: ScriptedTeamsGateway) -> TeamsReadingService:
    return build_teams_reading(fake_runtime(), teams_rig(gateway).driver)


def test_typeahead_rows_are_typed() -> None:
    results = parse_results(typed("atlas", *TYPEAHEAD), TODAY)

    assert results == [
        SearchResult(
            ResultType.GROUP_CHAT,
            "Project Atlas",
            chat="Project Atlas",
            members="Jane Doe, Omar Haddad",
        ),
        SearchResult(ResultType.SUGGESTION, "atlas budget"),
        SearchResult(ResultType.FILE, "Atlas plan.pptx, Created by Jane Doe"),
        SearchResult(
            ResultType.MESSAGE,
            "Omar Haddad: the atlas numbers are in",
            chat="Omar Haddad",
            time="Yesterday",
        ),
        SearchResult(ResultType.PERSON, "Alex Example"),
    ]
    assert results[0].to_record() == {
        "type": "group_chat",
        "text": "Project Atlas",
        "chat": "Project Atlas",
        "members": "Jane Doe, Omar Haddad",
        "time": "",
    }


def test_message_rows_take_chat_members_and_time_in_order() -> None:
    results = parse_results(
        typed(
            "atlas",
            "Message from Jane Doe at Jane Doe: draft attached",
            "Project Atlas",
            "Jane Doe, Omar Haddad",
            "10/19",
        ),
        TODAY,
    )

    assert results == [
        SearchResult(
            ResultType.MESSAGE,
            "Jane Doe: draft attached",
            chat="Project Atlas",
            members="Jane Doe, Omar Haddad",
            time="10/19",
        )
    ]


def test_no_popup_means_no_results() -> None:
    assert parse_results(idle(), TODAY) == []


def test_search_types_the_query_reads_results_and_clears_the_box() -> None:
    gateway = scripted(*TYPEAHEAD)

    results = reading(gateway).search("  atlas  ", 3)

    assert [result.type for result in results] == [
        ResultType.GROUP_CHAT,
        ResultType.SUGGESTION,
        ResultType.FILE,
    ]
    assert gateway.values == [("search-box", "atlas"), ("search-box", "")]
    assert gateway.keys == [Shortcut(Key.ESCAPE), Shortcut(Key.ESCAPE)]
    assert gateway.focused == ["search-box"]
    assert gateway.pressed == []


def test_search_without_suggestions_returns_empty_and_still_clears() -> None:
    gateway = scripted()

    assert reading(gateway).search("zzqx", 15) == []
    assert gateway.values[-1] == ("search-box", "")
    assert gateway.keys == [Shortcut(Key.ESCAPE), Shortcut(Key.ESCAPE)]


@pytest.mark.parametrize("query", ["a", " ", "x" * 201])
def test_bad_query_is_refused_before_any_ui(query: str) -> None:
    gateway = scripted(*TYPEAHEAD)

    with pytest.raises(TargetError, match="2 to 200 characters"):
        reading(gateway).search(query, 15)

    assert gateway.snapshots == 0
    assert gateway.values == []
