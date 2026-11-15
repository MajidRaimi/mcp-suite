from __future__ import annotations

from fakes.p6_trees import message_with_content
from fakes.teams_gateway import FakeTeamsGateway
from fakes.trees import chat_row, static, teams_menu, teams_window
from office_mcp.shared.node import Node

OMAR = "19:omar@unq.gbl.spaces"
ALICE = "19:alice@unq.gbl.spaces"
MINE_ID = "1790000000000"
THEIRS_ID = "1790000001000"
NEWEST_THEIRS_ID = "1790000002000"
MARK_UNREAD_ITEM = "item:Mark as unread"


def text_content(message_id: str, text: str) -> Node:
    return Node(role="AXGroup", dom_id=f"content-{message_id}", children=(static(text),))


def omar_messages(only_mine: bool = False) -> list[Node]:
    nodes = message_with_content(
        MINE_ID, "Jane Doe", text_content(MINE_ID, "draft sent"), mine=True
    )
    if only_mine:
        return nodes
    return [
        *nodes,
        *message_with_content(THEIRS_ID, "Omar Haddad", text_content(THEIRS_ID, "got it")),
        *message_with_content(
            NEWEST_THEIRS_ID, "Omar Haddad", text_content(NEWEST_THEIRS_ID, "one question")
        ),
    ]


class UnreadDesk:
    def __init__(
        self,
        gateway: FakeTeamsGateway,
        omar_unread: bool = True,
        only_mine: bool = False,
        honors_mark: bool = True,
    ) -> None:
        self.gateway = gateway
        self.open = ALICE
        self.omar_unread = omar_unread
        self.only_mine = only_mine
        self.honors_mark = honors_mark
        gateway.on_press = self.press
        gateway.menus = {
            f"body:{message_id}": teams_menu("Reply with quote", "Mark as unread")
            for message_id in (MINE_ID, THEIRS_ID, NEWEST_THEIRS_ID)
        }
        gateway.show(self.render())

    def rows(self) -> list[Node]:
        return [
            chat_row(OMAR, "Omar Haddad", "one question", "4:09 AM", unread=self.omar_unread),
            chat_row(ALICE, "Alice Example", "thanks", "10/19"),
        ]

    def render(self) -> Node:
        if self.open == OMAR:
            return teams_window(
                "Chat | Omar Haddad", self.rows(), omar_messages(self.only_mine), thread=OMAR
            )
        return teams_window("Chat | Alice Example", self.rows(), [], thread=ALICE)

    def press(self, node: Node) -> None:
        if node.handle == OMAR:
            self.open = OMAR
            self.omar_unread = False
        elif node.handle == MARK_UNREAD_ITEM and self.honors_mark:
            self.omar_unread = True
        self.gateway.show(self.render())
