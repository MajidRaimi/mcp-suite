from __future__ import annotations

import re

import pytest
from conftest import TODAY
from fakes.clock import fake_waiter
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.trees import folder_row, load_more_row, outlook_row, outlook_rows

from office_mcp.features.outlook.listing import parse_rows
from office_mcp.features.outlook.models import Row
from office_mcp.features.outlook.rows import (
    FolderLocator,
    RefRegistry,
    RowFinder,
    SearchLocator,
    SearchScope,
    is_key,
)
from office_mcp.features.outlook.window import FocusedTab, OutlookMode
from office_mcp.shared.errors import AmbiguousTargetError, TargetError, TargetNotFoundError

BUDGET = "  Alice Example, Project Atlas budget,     9:15 AM,        Numbers attached"
REVIEW = "  Omar Haddad, Atlas review,     Yesterday,        See you then"
TIMELINE = "  Alex Example, Atlas timeline,     Mon,        Timeline below"
FOLDERS = [
    folder_row("jane@example.com", level=0),
    folder_row("Inbox; 2 unread messages", handle="inbox"),
    folder_row("Archive", handle="archive"),
]


def finder(gateway: FakeOutlookGateway, registry: RefRegistry | None = None) -> RowFinder:
    return RowFinder(gateway, fake_waiter(), lambda: TODAY, registry or RefRegistry())


def keyed(*descriptions: str) -> list[Row]:
    return parse_rows(outlook_rows(*descriptions), TODAY)


@pytest.mark.parametrize(
    ("ref", "valid"),
    [("0123456789ab", True), ("0123456789AB", False), ("12", False), ("budget", False)],
)
def test_keys_are_twelve_lowercase_hex(ref: str, valid: bool) -> None:
    assert is_key(ref) is valid


def test_registry_evicts_the_least_recently_used_key() -> None:
    registry = RefRegistry(capacity=2)
    first, second, third = keyed(BUDGET, REVIEW, TIMELINE)
    inbox = FolderLocator("Inbox")

    registry.remember([first, second], inbox)
    assert registry.locate(first.summary.key) == inbox
    registry.remember([third], FolderLocator("Archive"))

    assert registry.locate(second.summary.key) is None
    assert registry.locate(first.summary.key) == inbox
    assert registry.locate(third.summary.key) == FolderLocator("Archive")


def test_rows_open_the_folder_label_rows_and_remember_keys() -> None:
    gateway = FakeOutlookGateway(rows=[], folders=FOLDERS, folder="Archive")
    gateway.folder_rows_by_name["Inbox"] = outlook_rows(BUDGET, REVIEW)
    registry = RefRegistry()

    rows = finder(gateway, registry).rows(FolderLocator("Inbox"))

    assert gateway.visible_calls() == ["select:inbox"]
    assert [row.summary.folder for row in rows] == ["Inbox", "Inbox"]
    assert [row.summary.view for row in rows] == ["focused", "focused"]
    assert registry.locate(rows[0].summary.key) == FolderLocator("Inbox")


def test_rows_leave_search_and_switch_the_tab_when_asked() -> None:
    gateway = FakeOutlookGateway(rows=outlook_rows(BUDGET), folders=FOLDERS)
    gateway.mode_value = OutlookMode.SEARCH
    gateway.tab_rows[FocusedTab.OTHER] = outlook_rows(TIMELINE)

    rows = finder(gateway).rows(FolderLocator("Inbox", view=FocusedTab.OTHER))

    assert gateway.visible_calls() == ["cancel_search", "switch_tab:other"]
    assert [row.summary.title for row in rows] == ["Alex Example, Atlas timeline"]
    assert rows[0].summary.view == "other"


def test_search_locator_runs_the_search_once() -> None:
    gateway = FakeOutlookGateway()
    gateway.search_rows = outlook_rows("Top Results, Expanded", BUDGET)
    rows_finder = finder(gateway)
    locator = SearchLocator("atlas", SearchScope.ALL)

    first = rows_finder.rows(locator)
    second = rows_finder.rows(locator)

    assert gateway.visible_calls() == ["run_search:atlas:All Mailboxes"]
    assert first[0].summary.group == "Top Results"
    assert [row.summary.key for row in second] == [row.summary.key for row in first]


def test_load_more_pages_are_pressed() -> None:
    gateway = FakeOutlookGateway(rows=[*outlook_rows(BUDGET), load_more_row()])
    gateway.pages = [outlook_rows(REVIEW, TIMELINE)]

    rows = finder(gateway).rows(FolderLocator("Inbox"), load_more=3)

    assert "press:load-more" in gateway.calls
    assert len(rows) == 3


@pytest.mark.parametrize("ref", ["0", "Project Atlas budget", "ABCDEF012345", "budget"])
def test_write_resolution_takes_keys_only(ref: str) -> None:
    rows = keyed(BUDGET, REVIEW)

    with pytest.raises(TargetError, match="use the key from outlook_conversations"):
        finder(FakeOutlookGateway()).resolve_for_write(ref, rows)


def test_write_resolution_rejects_missing_and_duplicated_keys() -> None:
    rows = keyed(BUDGET, REVIEW)
    rows_finder = finder(FakeOutlookGateway())
    twins = keyed(BUDGET, BUDGET)

    assert rows_finder.resolve_for_write(rows[1].summary.key, rows) is rows[1]
    with pytest.raises(TargetNotFoundError, match="changed or a day passed"):
        rows_finder.resolve_for_write("0" * 12, rows)
    with pytest.raises(AmbiguousTargetError) as raised:
        rows_finder.resolve_for_write(twins[0].summary.key, twins)
    assert len(raised.value.candidates) == 2


def test_a_conversation_in_top_results_and_all_results_resolves_to_the_listed_row() -> None:
    gateway = FakeOutlookGateway()
    gateway.search_rows = outlook_rows(
        "Top Results, Expanded", BUDGET, "All Results, Expanded", REVIEW, BUDGET
    )
    rows_finder = finder(gateway)
    locator = SearchLocator("atlas", SearchScope.ALL)
    rows = rows_finder.rows(locator)
    key = keyed(BUDGET)[0].summary.key

    for_read = rows_finder.resolve_for_read(key, rows)
    for_write = rows_finder.refresh(key, locator)

    assert for_read.summary.group == "All Results"
    assert for_write.summary.group == "All Results"


def test_the_same_key_twice_in_one_group_stays_ambiguous() -> None:
    rows = parse_rows(
        outlook_rows("Top Results, Expanded", BUDGET, "All Results, Expanded", BUDGET, BUDGET),
        TODAY,
    )

    with pytest.raises(AmbiguousTargetError):
        finder(FakeOutlookGateway()).resolve_for_write(rows[0].summary.key, rows)


def test_read_resolution_accepts_key_index_and_unique_text() -> None:
    rows = keyed(BUDGET, REVIEW, TIMELINE)
    rows_finder = finder(FakeOutlookGateway())

    assert rows_finder.resolve_for_read(rows[2].summary.key, rows) is rows[2]
    assert rows_finder.resolve_for_read(str(rows[1].summary.index), rows) is rows[1]
    assert rows_finder.resolve_for_read("budget", rows) is rows[0]
    with pytest.raises(AmbiguousTargetError, match="matches several conversations"):
        rows_finder.resolve_for_read("atlas", rows)
    with pytest.raises(TargetNotFoundError):
        rows_finder.resolve_for_read("nothing like it", rows)


def test_refresh_finds_the_key_again_and_loads_more_pages_if_needed() -> None:
    gateway = FakeOutlookGateway(rows=[*outlook_rows(BUDGET), load_more_row()])
    gateway.pages = [[outlook_row(REVIEW, handle="later")]]
    wanted = keyed(REVIEW)[0].summary.key

    row = finder(gateway).refresh(wanted, FolderLocator("Inbox"))

    assert row.node.handle == "later"
    assert row.summary.folder == "Inbox"


def test_refresh_errors_before_touching_the_ui_on_a_non_key() -> None:
    gateway = FakeOutlookGateway(rows=outlook_rows(BUDGET))

    with pytest.raises(TargetError):
        finder(gateway).refresh("budget", FolderLocator("Inbox"))
    assert gateway.calls == []


def test_refresh_reports_a_vanished_or_duplicated_key() -> None:
    gone = keyed(TIMELINE)[0].summary.key
    gateway = FakeOutlookGateway(rows=outlook_rows(BUDGET, REVIEW))
    with pytest.raises(TargetNotFoundError, match="list or search again"):
        finder(gateway).refresh(gone, FolderLocator("Inbox"))

    twin = keyed(BUDGET)[0].summary.key
    doubled = FakeOutlookGateway(rows=outlook_rows(BUDGET, BUDGET))
    with pytest.raises(AmbiguousTargetError):
        finder(doubled).refresh(twin, FolderLocator("Inbox"))


def test_locator_for_prefers_the_registry_then_the_fallback_then_the_screen() -> None:
    registry = RefRegistry()
    known = keyed(BUDGET)
    registry.remember(known, SearchLocator("atlas", SearchScope.MAILBOX))
    gateway = FakeOutlookGateway(folder="Archive", tab=FocusedTab.OTHER)
    rows_finder = finder(gateway, registry)

    assert rows_finder.locator_for(known[0].summary.key, "Inbox") == SearchLocator(
        "atlas", SearchScope.MAILBOX
    )
    assert rows_finder.locator_for("0" * 12, "Inbox") == FolderLocator("Inbox")
    assert rows_finder.locator_for("0" * 12, None) == FolderLocator("Archive", FocusedTab.OTHER)


def test_opening_an_unknown_folder_lists_the_known_ones() -> None:
    gateway = FakeOutlookGateway(folders=FOLDERS, folder="Archive")

    with pytest.raises(
        TargetNotFoundError, match=re.escape("Folders: jane@example.com, Inbox, Archive")
    ):
        finder(gateway).open_folder("Nowhere")
