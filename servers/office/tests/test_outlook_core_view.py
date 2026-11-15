from __future__ import annotations

import pytest
from fakes.clock import fake_waiter
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.trees import folder_row, outlook_rows

from office_mcp.features.outlook.rows import SearchScope, SearchState
from office_mcp.features.outlook.view import ViewState, capture_view, restore_view
from office_mcp.features.outlook.window import FocusedTab, OutlookMode
from office_mcp.shared.errors import CleanupError
from office_mcp.shared.node import Node

ROW = "  Alice Example, Project Atlas budget,     9:15 AM,        Numbers attached"
FOLDERS = [folder_row("Inbox; 2 unread messages", handle="inbox"), folder_row("Archive")]
INBOX_VIEW = ViewState(OutlookMode.MAIL, "Inbox", FocusedTab.FOCUSED, None, had_selection=False)


def sheet(*buttons: str) -> Node:
    kids = tuple(Node(role="AXButton", title=title, handle=f"sheet:{title}") for title in buttons)
    return Node(role="AXSheet", description="alert", children=kids)


def test_capture_reads_mode_folder_tab_search_and_selection() -> None:
    gateway = FakeOutlookGateway(rows=outlook_rows(ROW), folder="Archive", tab=FocusedTab.OTHER)
    gateway.selected = {"row0"}

    assert capture_view(gateway) == ViewState(
        OutlookMode.MAIL, "Archive", FocusedTab.OTHER, None, had_selection=True
    )


def test_capture_in_search_keeps_the_query_and_scope() -> None:
    gateway = FakeOutlookGateway()
    gateway.run_search("atlas", SearchScope.SUBFOLDERS)
    gateway.calls.clear()

    state = capture_view(gateway)

    assert state.mode is OutlookMode.SEARCH
    assert state.search == SearchState("atlas", SearchScope.SUBFOLDERS)
    assert state.folder == ""


def test_restore_runs_every_step_in_order() -> None:
    gateway = FakeOutlookGateway(
        rows=outlook_rows(ROW), folders=FOLDERS, folder="Archive", tab=FocusedTab.OTHER
    )
    gateway.popovers = 1
    gateway.suggestions = 1
    gateway.sheets = [sheet("Discard", "Cancel")]
    gateway.run_search("atlas", SearchScope.ALL)
    gateway.rows = outlook_rows(ROW)
    gateway.selected = {"row0"}
    gateway.calls.clear()

    notes = restore_view(gateway, INBOX_VIEW, fake_waiter())

    assert gateway.visible_calls() == [
        "dismiss_popovers:1",
        "close_suggestions",
        "press_button:Cancel",
        "cancel_search",
        "select:inbox",
        "switch_tab:focused",
        "deselect:row0",
    ]
    assert notes[0].startswith("closed 3 leftover")
    assert notes[-1] == "cleared the selection"


def test_restore_runs_the_owner_search_again() -> None:
    gateway = FakeOutlookGateway()
    before = ViewState(
        OutlookMode.SEARCH, "", FocusedTab.FOCUSED, SearchState("atlas", SearchScope.ALL), True
    )

    restore_view(gateway, before, fake_waiter())

    assert gateway.visible_calls() == ["run_search:atlas:All Mailboxes"]


def test_restore_goes_back_to_the_calendar_without_touching_mail_state() -> None:
    gateway = FakeOutlookGateway(rows=outlook_rows(ROW))
    gateway.selected = {"row0"}
    before = ViewState(OutlookMode.CALENDAR, "", None, None, had_selection=False)

    restore_view(gateway, before, fake_waiter())

    assert gateway.visible_calls() == ["show_module:Calendar"]


def test_restore_keeps_going_after_a_failed_step_and_reports_it() -> None:
    gateway = FakeOutlookGateway(rows=outlook_rows(ROW), tab=FocusedTab.OTHER)
    gateway.sheets = [sheet("Send Anyway")]
    gateway.selected = {"row0"}

    with pytest.raises(CleanupError, match="an Outlook sheet is still open"):
        restore_view(gateway, INBOX_VIEW, fake_waiter())

    assert "switch_tab:focused" in gateway.calls
    assert "deselect:row0" in gateway.calls
    assert not any("Send Anyway" in call for call in gateway.calls)


def test_capture_in_a_folder_search_keeps_the_searched_folder() -> None:
    gateway = FakeOutlookGateway(folders=FOLDERS, folder="Archive")
    gateway.run_search("budget", SearchScope.CURRENT_FOLDER)
    gateway.selected = {"archive"}
    gateway.folders = [FOLDERS[0], folder_row("Archive", level=0, handle="archive")]

    state = capture_view(gateway)

    assert state.folder == "Archive"


def test_restore_reopens_the_searched_folder_before_a_folder_search() -> None:
    gateway = FakeOutlookGateway(
        folders=[FOLDERS[0], folder_row("Archive", level=0, handle="archive")], folder="Inbox"
    )
    before = ViewState(
        OutlookMode.SEARCH,
        "Archive",
        FocusedTab.FOCUSED,
        SearchState("budget", SearchScope.CURRENT_FOLDER),
        True,
    )

    restore_view(gateway, before, fake_waiter())

    assert gateway.visible_calls() == ["select:archive", "run_search:budget:Current Folder"]
    assert gateway.folder == "Archive"
