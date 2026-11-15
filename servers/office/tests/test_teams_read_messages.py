from __future__ import annotations

from datetime import UTC

from fakes.p6_trees import message_with_content, quoted_content
from fakes.trees import message, static, teams_window

from office_mcp.features.teams.messages.models import Message, Quote
from office_mcp.features.teams.messages.parse import message_body, parse_messages
from office_mcp.shared.node import Node

QUOTED = "1790000005000"
PLAIN = "1790000009000"
QUOTE_TITLE = "Begin quote, Jane Doe, 9/21/2026, 2:13 PM, can we meet on Project Atlas?"
QUOTE_PIECES = ["Jane Doe", "9/21/2026 2:13 PM", "can we meet on Project Atlas?"]


def parsed(*nodes: Node) -> list[Message]:
    return parse_messages(teams_window("Chat | Omar Haddad", [], list(nodes)), UTC)


def only(messages: list[Message]) -> Message:
    assert len(messages) == 1
    return messages[0]


def test_quote_is_parsed_and_reply_text_kept_apart() -> None:
    content = quoted_content(QUOTED, QUOTE_TITLE, QUOTE_PIECES, "yes, at 3")

    found = only(parsed(*message_with_content(QUOTED, "Omar Haddad", content)))

    assert found.quote == Quote("Jane Doe", "9/21/2026 2:13 PM", "can we meet on Project Atlas?")
    assert found.text == "yes, at 3"
    assert found.author == "Omar Haddad"
    assert found.to_record()["quote"] == {
        "author": "Jane Doe",
        "sent": "9/21/2026 2:13 PM",
        "text": "can we meet on Project Atlas?",
    }


def test_reply_text_is_cut_from_a_content_title_that_holds_the_quote() -> None:
    content = quoted_content(
        QUOTED, QUOTE_TITLE, QUOTE_PIECES, "yes, at 3", title=f"{QUOTE_TITLE} yes, at 3"
    )

    found = only(parsed(*message_with_content(QUOTED, "Omar Haddad", content)))

    assert found.text == "yes, at 3"
    assert found.quote is not None and found.quote.author == "Jane Doe"


def test_quote_author_with_a_comma_and_24_hour_time() -> None:
    title = "Begin quote, Doe, Jane, 21/9/2026, 14:13, budget draft"
    content = quoted_content(QUOTED, title, [], "noted")

    found = only(parsed(*message_with_content(QUOTED, "Alex Example", content)))

    assert found.quote == Quote("Doe, Jane", "21/9/2026 14:13", "budget draft")
    assert found.text == "noted"


def test_quote_falls_back_to_its_pieces_when_the_title_is_unusual() -> None:
    content = quoted_content(QUOTED, "Begin quote, Jane Doe", QUOTE_PIECES, "sure")

    found = only(parsed(*message_with_content(QUOTED, "Alex Example", content)))

    assert found.quote == Quote("Jane Doe", "9/21/2026 2:13 PM", "can we meet on Project Atlas?")
    assert found.text == "sure"


def test_messages_without_a_quote_keep_quote_none_and_text() -> None:
    nodes = message(PLAIN, "Omar Haddad", "see you at 4", "October 19, 2026 9:00 AM")

    found = only(parsed(*nodes))

    assert found.quote is None
    assert found.edited is False
    assert found.text == "see you at 4"
    assert set(found.to_record()) == {"id", "author", "mine", "sent_at", "text", "quote", "edited"}


def test_edited_mark_next_to_the_header_sets_edited() -> None:
    content = Node(role="AXGroup", dom_id=f"content-{PLAIN}", children=(static("fixed typo"),))
    nodes = [
        *message_with_content(PLAIN, "Alex Example", content, mine=True, edited=True),
        *message("1790000019000", "Omar Haddad", "thanks", "October 19, 2026 9:05 AM"),
    ]

    first, second = parsed(*nodes)

    assert first.edited is True and first.mine is True
    assert first.text == "fixed typo"
    assert second.edited is False


def test_edited_mark_inside_the_body_sets_edited() -> None:
    content = Node(
        role="AXGroup",
        dom_id=f"content-{PLAIN}",
        children=(static("updated plan"), Node(role="AXGroup", dom_id=f"edited-{PLAIN}")),
    )

    found = only(parsed(*message_with_content(PLAIN, "Alex Example", content)))

    assert found.edited is True


def test_edited_mark_after_the_body_belongs_to_that_message() -> None:
    later = "1790000010000"
    first_body = Node(
        role="AXGroup",
        dom_id=f"message-body-{PLAIN}",
        classes=("fui-ChatMessage__body",),
        children=(static("first draft of the plan"),),
    )
    second_body = Node(
        role="AXGroup",
        dom_id=f"message-body-{later}",
        classes=("fui-ChatMessage__body",),
        children=(static("and the dates"),),
    )

    first, second = parsed(
        Node(role="AXHeading", title="preview by Jane Doe", value="4"),
        first_body,
        Node(role="AXStaticText", value="Edited"),
        second_body,
    )

    assert first.edited is True
    assert second.edited is False
    assert second.author == "Jane Doe"


def test_message_body_finds_the_body_by_id() -> None:
    root = teams_window(
        "Chat | Omar Haddad", [], message(PLAIN, "Omar Haddad", "hi", "October 19, 2026 9:00 AM")
    )

    body = message_body(root, PLAIN)

    assert body is not None and body.dom_id == f"message-body-{PLAIN}"
    assert message_body(root, "missing") is None
