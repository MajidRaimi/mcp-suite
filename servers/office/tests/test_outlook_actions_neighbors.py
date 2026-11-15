from __future__ import annotations

from fakes.drivers import FAKE_TODAY
from fakes.p2_outlook import Mail, atlas, default_folders, lunch, offsite

from office_mcp.features.outlook.actions.neighbors import adjacent_unread, turned_read
from office_mcp.features.outlook.actions.service import is_protected_folder
from office_mcp.features.outlook.listing import parse_folders, parse_rows
from office_mcp.features.outlook.models import Row
from office_mcp.shared.node import Node

FOLDERS = parse_folders(default_folders())


def rows(*mails: Mail) -> list[Row]:
    nodes = [
        Node(role="AXRow", description=mail.description(), handle=f"m{index}")
        for index, mail in enumerate(mails)
    ]
    return parse_rows(nodes, FAKE_TODAY)


def test_adjacent_unread_lists_only_unread_rows_above_and_below() -> None:
    listed = rows(lunch(unread=True), atlas(), offsite(unread=False))

    assert adjacent_unread(listed, listed[1].summary.key) == (listed[0].summary.key,)
    assert adjacent_unread(listed, listed[0].summary.key) == ()
    assert adjacent_unread(listed, "0123456789ab") == ()


def test_turned_read_reports_watched_rows_that_lost_unread() -> None:
    before = rows(lunch(unread=True), offsite(unread=True))
    after = rows(lunch(unread=False), offsite(unread=True))
    watched = tuple(row.summary.key for row in before)

    assert [row.summary.key for row in turned_read(watched, after)] == [after[0].summary.key]


def test_protected_folders_include_paths_and_subfolders() -> None:
    assert is_protected_folder("Deleted Items", FOLDERS)
    assert is_protected_folder("Old", FOLDERS)
    assert is_protected_folder("jane@example.com/Recoverable Items/Purges", [])
    assert not is_protected_folder("Inbox", FOLDERS)
    assert not is_protected_folder("Travel", FOLDERS)
