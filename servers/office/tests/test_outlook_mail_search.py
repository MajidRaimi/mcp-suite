from __future__ import annotations

import pytest
from fakes.p1_mail import MailGateway, PaneSpec, mail_rig, mail_row, reading_pane, section
from fakes.trees import folder_row, outlook_rows

from office_mcp.features.outlook.mail.search import SearchArea
from office_mcp.features.outlook.rows import SearchScope, SearchState
from office_mcp.features.outlook.window import OutlookMode
from office_mcp.shared.errors import UiTimeoutError
from office_mcp.shared.node import Node

TOP = mail_row(
    "Alice Example", "Project Atlas budget", "9:15 AM", "Numbers attached", folder="Inbox"
)
ARCHIVED = mail_row(
    "Omar Haddad", "Atlas review", "Yesterday", "See you then", folder="Archive", unread=True
)
SEARCH_ROWS = outlook_rows(section("Top Results"), TOP, section("All Results"), ARCHIVED)


def search_gateway() -> MailGateway:
    gateway = MailGateway(
        rows=outlook_rows(mail_row("Bob Sample", "Lunch", "8:00 AM", "Noon?")),
        folders=[folder_row("Inbox; 1 unread message", handle="inbox"), folder_row("Archive")],
    )
    gateway.search_rows = SEARCH_ROWS
    return gateway


def test_search_parses_sections_and_folders_then_leaves_search() -> None:
    gateway = search_gateway()
    rig = mail_rig(gateway)

    result = rig.mail.search("atlas", SearchArea.MAILBOX, None, 25)
    record = result.to_record()

    assert record["query"] == "atlas"
    assert record["scope"] == "mailbox"
    assert record["count"] == 2
    first, second = record["results"]
    assert (first["group"], first["folder"], first["preview"]) == (
        "Top Results",
        "Inbox",
        "Numbers attached",
    )
    assert (second["group"], second["folder"], second["unread"]) == (
        "All Results",
        "Archive",
        True,
    )
    assert "run_search:atlas:Current Mailbox" in gateway.calls
    assert gateway.mode_value is OutlookMode.MAIL
    assert gateway.search is None


def test_search_keys_resolve_by_running_the_search_again() -> None:
    gateway = search_gateway()
    gateway.panes["row3"] = reading_pane(PaneSpec(subject="Atlas review"))
    rig = mail_rig(gateway)
    key = rig.mail.search("atlas", SearchArea.ALL, None, 25).results[1].key

    email = rig.mail.read(key, None)

    assert email.summary is not None and email.summary.folder == "Archive"
    assert gateway.calls.count("run_search:atlas:All Mailboxes") == 2
    assert gateway.search is None
    assert gateway.is_unread("row3")


def test_search_puts_the_previous_scope_back_before_leaving() -> None:
    gateway = search_gateway()
    rig = mail_rig(gateway, scope=SearchScope.MAILBOX)

    rig.mail.search("atlas", SearchArea.ALL, None, 25)

    calls = gateway.visible_calls()
    assert rig.control.chosen == [SearchScope.MAILBOX]
    assert calls.index("choose_scope:Current Mailbox") < calls.index("cancel_search")


def test_search_with_the_same_scope_changes_nothing_back() -> None:
    rig = mail_rig(search_gateway(), scope=SearchScope.MAILBOX)

    rig.mail.search("atlas", SearchArea.MAILBOX, None, 25)

    assert rig.control.chosen == []


def test_zero_results_is_a_count_of_zero() -> None:
    gateway = search_gateway()
    gateway.search_rows = []
    rig = mail_rig(gateway)

    record = rig.mail.search("nothing like this", SearchArea.MAILBOX, None, 25).to_record()

    assert record["count"] == 0
    assert record["results"] == []
    assert gateway.search is None


def test_folder_scope_opens_the_folder_first() -> None:
    gateway = search_gateway()
    rig = mail_rig(gateway)

    rig.mail.search("atlas", SearchArea.FOLDER, "Archive", 25)

    calls = gateway.visible_calls()
    assert calls.index("select:Archive") < calls.index("run_search:atlas:Current Folder")
    assert gateway.folder == "Inbox"


def test_limit_trims_results() -> None:
    rig = mail_rig(search_gateway())

    assert rig.mail.search("atlas", SearchArea.MAILBOX, None, 1).to_record()["count"] == 1


class FailingSearchGateway(MailGateway):
    def message_rows(self) -> list[Node]:
        if self.search is not None:
            raise UiTimeoutError("the result list vanished")
        return super().message_rows()


def test_search_is_cancelled_even_when_listing_fails() -> None:
    gateway = FailingSearchGateway(folders=[folder_row("Inbox", handle="inbox")])
    rig = mail_rig(gateway)

    with pytest.raises(UiTimeoutError):
        rig.mail.search("atlas", SearchArea.ALL, None, 25)
    assert gateway.search is None
    assert rig.control.chosen == [SearchScope.MAILBOX]


def test_owner_search_is_run_again_afterwards() -> None:
    gateway = search_gateway()
    gateway.search = SearchState("lunch", SearchScope.CURRENT_FOLDER)
    gateway.mode_value = OutlookMode.SEARCH
    rig = mail_rig(gateway)

    rig.mail.search("atlas", SearchArea.MAILBOX, None, 25)

    assert gateway.search == SearchState("lunch", SearchScope.CURRENT_FOLDER)
    assert gateway.visible_calls()[-1] == "run_search:lunch:Current Folder"


def test_search_from_the_calendar_switches_to_mail_and_back() -> None:
    gateway = search_gateway()
    gateway.mode_value = OutlookMode.CALENDAR
    rig = mail_rig(gateway)

    rig.mail.search("atlas", SearchArea.MAILBOX, None, 25)

    calls = gateway.visible_calls()
    assert calls.index("show_module:Mail") < calls.index("run_search:atlas:Current Mailbox")
    assert gateway.mode_value is OutlookMode.CALENDAR
