from __future__ import annotations

from collections.abc import Callable
from datetime import date

from office_mcp.features.teams.chats.models import Chat, ChatRow
from office_mcp.features.teams.chats.parse import parse_chats
from office_mcp.features.teams.gateway import TeamsGateway
from office_mcp.shared.errors import OfficeError
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize

CHAT_APP_PREFIX = "Chat ("
CHAT_APP_ROLE = "AXCheckBox"
MAX_CANDIDATES = 8


class ChatService:
    def __init__(self, gateway: TeamsGateway, today: Callable[[], date]) -> None:
        self._gateway = gateway
        self._today = today

    def rows(self) -> list[ChatRow]:
        root = self._gateway.snapshot()
        rows = parse_chats(root, self._today())
        if rows:
            return rows
        switch = root.find(is_chat_app_button)
        if switch is None:
            raise OfficeError("no chat list found in the Teams window")
        self._gateway.press(switch)
        return parse_chats(self._gateway.snapshot(), self._today())

    def filtered(self, unread_only: bool, section: str | None, limit: int) -> list[Chat]:
        chats = [row.chat for row in self.rows()]
        if unread_only:
            chats = [chat for chat in chats if chat.unread]
        if section:
            wanted = normalize(section)
            chats = [chat for chat in chats if wanted in normalize(chat.section)]
        return chats[:limit] if limit > 0 else chats

    def resolve(self, query: str, rows: list[ChatRow]) -> ChatRow:
        exact = [row for row in rows if row.chat.thread_id == query]
        if exact:
            return exact[0]
        return pick(query, rows)


def is_chat_app_button(node: Node) -> bool:
    return node.role == CHAT_APP_ROLE and node.title.startswith(CHAT_APP_PREFIX)


def pick(query: str, rows: list[ChatRow]) -> ChatRow:
    wanted = normalize(query)
    if not wanted:
        raise OfficeError("empty chat name")
    names = [(row, normalize(row.chat.name)) for row in rows]
    for matches in (
        [row for row, name in names if name == wanted],
        [row for row, name in names if name.startswith(wanted)],
        [row for row, name in names if wanted in name],
        [row for row, name in names if all(word in name for word in wanted.split())],
    ):
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            listed = ", ".join(row.chat.name for row in matches[:MAX_CANDIDATES])
            raise OfficeError(f"{query!r} matches several chats: {listed}. Be more specific.")
    raise OfficeError(f"no chat named {query!r} in the chat list")
