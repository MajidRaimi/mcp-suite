from __future__ import annotations

import pytest
from conftest import TODAY
from fakes.trees import folder_row

from office_mcp.features.outlook.listing import (
    parse_folder,
    parse_folders,
    parse_rows,
    pick_folder_row,
    row_flags,
)
from office_mcp.features.outlook.models import Folder
from office_mcp.shared.node import Node

FOLDER_PANE = [
    folder_row("Favorites", level=0),
    folder_row("Inbox", level=1, handle="fav-inbox"),
    folder_row("Sent Items", level=1, handle="fav-sent"),
    folder_row("jane@example.com", level=0),
    folder_row("Inbox; 206 unread messages", level=1, handle="inbox"),
    folder_row("Drafts; 3 messages", level=1, handle="drafts"),
    folder_row("Groups", level=1, handle="groups"),
    folder_row("Project Atlas; 10 unseen messages", level=2, handle="atlas"),
    folder_row("Archive", level=1, handle="archive"),
    folder_row("Saved Searches", level=0),
    folder_row("Unread Mails", level=1, handle="unread-search"),
]


@pytest.mark.parametrize(
    ("description", "expected"),
    [
        ("Inbox; 12 unread messages", Folder("Inbox", 12, path="Inbox")),
        ("Projects; 1 unread message", Folder("Projects", 1, path="Projects")),
        ("Drafts; 3 messages", Folder("Drafts", 0, total=3, path="Drafts")),
        (
            "Project Atlas; 10 unseen messages",
            Folder("Project Atlas", 0, unseen=10, path="Project Atlas"),
        ),
        ("Archive", Folder("Archive", 0, path="Archive")),
    ],
)
def test_parses_the_three_count_suffixes(description: str, expected: Folder) -> None:
    assert parse_folder(description) == expected


def test_paths_follow_levels_and_favorites_are_marked() -> None:
    folders = parse_folders(FOLDER_PANE)

    by_path = {folder.path: folder for folder in folders}
    assert by_path["Favorites/Inbox"].favorite
    assert by_path["Favorites/Sent Items"].favorite
    assert not by_path["jane@example.com/Inbox"].favorite
    assert by_path["jane@example.com/Inbox"].unread == 206
    assert by_path["jane@example.com/Groups/Project Atlas"].unseen == 10
    assert by_path["jane@example.com/Groups/Project Atlas"].level == 2
    assert by_path["jane@example.com/Archive"].level == 1
    assert by_path["Saved Searches/Unread Mails"].path == "Saved Searches/Unread Mails"


def test_picking_a_folder_prefers_the_account_row_over_the_favorite() -> None:
    assert pick_folder_row(FOLDER_PANE, "inbox") is FOLDER_PANE[4]
    assert pick_folder_row(FOLDER_PANE, "Sent Items") is FOLDER_PANE[2]
    assert pick_folder_row(FOLDER_PANE, "jane@example.com/Groups/Project Atlas") is FOLDER_PANE[7]
    assert pick_folder_row(FOLDER_PANE, "proj") is FOLDER_PANE[7]
    assert pick_folder_row(FOLDER_PANE, "Nowhere") is None


SEARCH_ROWS = [
    "Top Results, Expanded",
    "  Alice Example, Project Atlas budget,     Yesterday,  Folder: Inbox,  Has files,   Numbers attached",
    "All Results, Expanded",
    " Unread, Flagged,  Omar Haddad, Atlas review,  Replied  18/10/2026,  Folder: Sent Items,   See you then",
    "Pinned, Expanded",
    "  Jane Doe, Pinned note,     Mon,      Keep this",
]


def test_search_rows_take_the_section_folder_token_and_flags() -> None:
    nodes = [Node(role="AXRow", description=text) for text in SEARCH_ROWS]

    budget, review, note = (row.summary for row in parse_rows(nodes, TODAY))

    assert (budget.group, budget.folder, budget.preview) == (
        "Top Results",
        "Inbox",
        "Numbers attached",
    )
    assert budget.has_attachments and not budget.flagged
    assert (review.group, review.folder, review.preview) == (
        "All Results",
        "Sent Items",
        "See you then",
    )
    assert review.flagged and review.replied and review.unread
    assert note.group == "Pinned" and note.pinned


def test_row_flags_map_known_flags_to_booleans() -> None:
    assert row_flags(("flagged", "has attachment", "forwarded")) == {
        "flagged": True,
        "replied": False,
        "forwarded": True,
        "has_attachments": True,
        "pinned": False,
    }


COLLAPSED = "  2 messages, Omar Haddad, Bob Sample, Atlas review, Q3,     9:15 AM,     See you then"
EXPANDED = [
    "  1 unread message, Expanded,  2 messages,  Atlas review, Q3,     9:15 AM,         ",
    " Unread, Omar  Haddad, Atlas review, Q3,   9:15 AM,        See you then",
    "  Bob Sample, Atlas review, Q3,   8:50 AM, Folder: Sent Items,       Earlier note",
]


def test_an_expanded_conversation_keeps_the_key_it_had_collapsed() -> None:
    collapsed = parse_rows([Node(role="AXRow", description=COLLAPSED)], TODAY)
    expanded = parse_rows([Node(role="AXRow", description=text) for text in EXPANDED], TODAY)

    assert len(expanded) == 1
    header = expanded[0].summary
    assert header.key == collapsed[0].summary.key
    assert header.preview == "See you then"
    assert header.unread
    assert len(header.thread) == 2


def test_a_flag_or_reaction_after_the_time_does_not_change_the_key() -> None:
    plain = "  Omar Haddad, Atlas review,     9:15 AM,        See you then"
    flagged = "  Omar Haddad, Atlas review,     9:15 AM,  Flagged,  Has reactions,   See you then"

    before, after = (
        parse_rows([Node(role="AXRow", description=text)], TODAY)[0].summary
        for text in (plain, flagged)
    )

    assert after.key == before.key
    assert after.flagged
    assert "has reactions" in after.flags
    assert after.preview == "See you then"


def test_a_completed_mark_without_a_comma_is_a_flag_not_preview() -> None:
    plain = "  Omar Haddad, Atlas review,     9:15 AM,        See you then"
    done = "  Omar Haddad, Atlas review,     9:15 AM,  Marked as completed       See you then"

    before, after = (
        parse_rows([Node(role="AXRow", description=text)], TODAY)[0].summary
        for text in (plain, done)
    )

    assert after.key == before.key
    assert "completed" in after.flags
    assert not after.flagged
