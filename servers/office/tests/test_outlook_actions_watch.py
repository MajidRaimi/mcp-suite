from __future__ import annotations

from conftest import TODAY
from fakes.clock import fake_waiter
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.trees import outlook_rows

from office_mcp.features.outlook.actions.watch import RowWatch, is_flagged, is_unflagged

EXPANDED = (
    "  1 unread message, Expanded,  2 messages,  Atlas review, Q3,     9:15 AM,         ",
    " Unread, Omar  Haddad, Atlas review, Q3,   9:15 AM,        See you then",
    "  Bob Sample, Atlas review, Q3,   8:50 AM, Folder: Sent Items,       Earlier note",
)


def test_a_selected_expanded_conversation_keeps_its_listed_key() -> None:
    gateway = FakeOutlookGateway(rows=outlook_rows(*EXPANDED))
    gateway.selected = {"row0"}
    watch = RowWatch(gateway, fake_waiter(), lambda: TODAY)

    listed = watch.current()[0].summary.key

    assert watch.selected_keys() == [listed]
    assert watch.wait_selected(listed)


def test_a_completed_flag_is_not_counted_as_cleared() -> None:
    gateway = FakeOutlookGateway(
        rows=outlook_rows("  Omar Haddad, Atlas review,     9:15 AM,  Marked as completed   Hi")
    )
    summary = RowWatch(gateway, fake_waiter(), lambda: TODAY).current()[0].summary

    assert not is_flagged(summary)
    assert not is_unflagged(summary)
