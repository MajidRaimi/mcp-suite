from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field

from fakes.ax_backend import FakeElement, element
from fakes.pasteboard import FakePasteboard
from fakes.trees import chat_header, chat_row, message, rail
from office_mcp.shared.errors import ElementNotFoundError
from office_mcp.shared.events import Key, Modifier, Shortcut
from office_mcp.shared.node import Node

PLACEHOLDER = "Type a message"
OWNER = "Alex Example"
STAMP = "October 20, 2026 9:00 AM"
TIME_LABEL = "9:00 AM"
PASTE = Shortcut(Key.V, Modifier.COMMAND)
SEND_KEYS = Shortcut(Key.RETURN, Modifier.COMMAND)
BACKSPACE = Shortcut(Key.DELETE)
SEND_HANDLE = "send-button"
CONFIRM_HANDLE = "confirm-delete"
CANCEL_HANDLE = "cancel-delete"
REACTIONS = ("Like", "Heart", "Laugh", "Surprised")
FIRST_ID = 1790000000000
DELETED_TEXT = "This message has been deleted."
BODY_PREFIX = "message-body-"
MINE_MENU = ("Reply with quote", "Edit", "Delete this message", "Copy message link")
DELETE_TITLES = frozenset({"Delete", "Delete this message"})
OTHERS_MENU = ("Reply with quote", "Forward", "Mark as unread", "Pin for everyone")


@dataclass
class FakeMessage:
    id: str
    author: str
    text: str
    mine: bool = False
    deleted: bool = False
    reactions: set[str] = field(default_factory=set)

    def shown_text(self) -> str:
        return DELETED_TEXT if self.deleted else self.text


@dataclass
class FakeChat:
    thread: str
    name: str
    kind: str = "Chat"
    participants: list[str] = field(default_factory=list)
    messages: list[FakeMessage] = field(default_factory=list)
    older: list[FakeMessage] = field(default_factory=list)
    unread: bool = False
    draft: str = ""
    quote: FakeMessage | None = None
    quote_text: str = ""
    listed: bool = True
    box: FakeElement = field(init=False)

    def __post_init__(self) -> None:
        self.box = element(
            "AXTextArea", AXValue=self.draft or PLACEHOLDER + "\n", AXTitle=PLACEHOLDER
        )


def own(message_id: int, text: str) -> FakeMessage:
    return FakeMessage(str(message_id), OWNER, text, mine=True)


def other(message_id: int, author: str, text: str) -> FakeMessage:
    return FakeMessage(str(message_id), author, text)


@dataclass
class WorldOptions:
    send_works: bool = True
    posted_text: str | None = None
    send_on_newline: bool = False
    popover: bool = True
    reaction_flips: bool = True
    reaction_state_visible: bool = True
    quote_in_value: bool = True
    quote_works: bool = True
    delete_confirm: bool = False
    delete_works: bool = True
    unread_works: bool = True
    send_button: bool = True
    substitutions: dict[str, str] = field(default_factory=dict)
    selected_text_works: bool = False
    menu_items: dict[str, tuple[str, ...]] = field(default_factory=dict)


class TeamsWorld:
    def __init__(
        self,
        chats: list[FakeChat],
        open_thread: str | None = None,
        options: WorldOptions | None = None,
    ) -> None:
        self.chats = chats
        self.open_thread = open_thread
        self.options = options or WorldOptions()
        self.pasteboard = FakePasteboard([{"public.utf8-plain-text": b"owner clipboard"}])
        self.focused_message: str | None = None
        self.dialog: str | None = None
        self.pressed: list[object] = []
        self.focused: list[str] = []
        self.menus_opened: list[str] = []
        self.scrolled: list[str] = []
        self.keys: list[Shortcut] = []
        self.commits = 0
        self.snapshots = 0
        self.close_calls = 0

    def chat(self, thread: str) -> FakeChat:
        return next(chat for chat in self.chats if chat.thread == thread)

    def opened(self) -> FakeChat | None:
        if self.open_thread is None:
            return None
        return self.chat(self.open_thread)

    def box_text(self, thread: str) -> str:
        value = str(self.chat(thread).box.attributes.get("AXValue", ""))
        return "" if value.rstrip("\n") == PLACEHOLDER else value

    def set_box(self, chat: FakeChat, text: str) -> None:
        chat.box.attributes["AXValue"] = text or PLACEHOLDER + "\n"
        chat.box.attributes["AXSelectedTextRange"] = (len(text), 0)

    def snapshot(self) -> Node:
        self.snapshots += 1
        chat = self.opened()
        children: list[Node] = [rail(), Node(role="AXGroup", children=tuple(self._rows()))]
        if chat is not None:
            children.append(chat_header(chat.thread, chat.participants))
            children.append(
                Node(role="AXGroup", dom_id="chat-pane-list", children=tuple(self._pane(chat)))
            )
            children.append(self._compose(chat))
        if self.dialog is not None:
            children.append(self._dialog())
        title = f"Chat | {chat.name}" if chat is not None else "Chat"
        return Node(role="AXWindow", title=title, children=tuple(children))

    def press(self, node: Node) -> None:
        handle = node.handle
        self.pressed.append(handle)
        if isinstance(handle, str) and any(chat.thread == handle for chat in self.chats):
            self.open_thread = handle
            self.focused_message = None
        elif handle == SEND_HANDLE:
            self._commit()
        elif handle == CONFIRM_HANDLE:
            self._delete_confirmed()
        elif isinstance(handle, tuple) and handle[0] == "toggle":
            self._toggle(str(handle[1]), str(handle[2]))
        elif isinstance(handle, tuple) and handle[0] == "item":
            self._item(str(handle[1]), str(handle[2]))

    def scroll_into_view(self, node: Node) -> None:
        self.scrolled.append(node.dom_id)
        chat = self.opened()
        if chat is not None and chat.older:
            chat.messages[:0] = chat.older
            chat.older = []

    def focus(self, node: Node) -> None:
        self.focused.append(node.dom_id)
        if node.dom_id.startswith(BODY_PREFIX):
            self.focused_message = node.dom_id.removeprefix(BODY_PREFIX)

    def set_value(self, node: Node, text: str) -> None:
        raise AssertionError("the compose box must never be written with a raw AXValue")

    def show_menu(self, node: Node) -> Node:
        message_id = node.dom_id.removeprefix(BODY_PREFIX)
        found = self._message(message_id)
        if found is None:
            raise ElementNotFoundError(f"no menu for {node.dom_id}")
        self.menus_opened.append(message_id)
        titles = self.options.menu_items.get(message_id, MINE_MENU if found.mine else OTHERS_MENU)
        toggles = tuple(self._toggle_node(found, name) for name in REACTIONS)
        items = tuple(
            Node(role="AXMenuItem", title=title, handle=("item", message_id, title))
            for title in titles
        )
        return Node(role="AXMenu", children=(*toggles, *items), handle="menu")

    def close_popups(self) -> int:
        self.close_calls += 1
        closed = 1 if self.dialog is not None else 0
        self.dialog = None
        return closed

    def press_key(self, shortcut: Shortcut) -> None:
        self.keys.append(shortcut)
        self.on_event(0, shortcut)

    def element(self, node: Node) -> object:
        if node.handle is None:
            raise ElementNotFoundError("no handle")
        return node.handle

    def on_event(self, _pid: int, shortcut: Shortcut) -> None:
        chat = self.opened()
        if chat is None:
            return
        if shortcut == PASTE:
            self._paste(chat, self.pasteboard.text() or "")
        elif shortcut == BACKSPACE:
            self._backspace(chat)
        elif shortcut == SEND_KEYS:
            self._commit()

    def install_box_hooks(self) -> None:
        if not self.options.selected_text_works:
            return
        for chat in self.chats:
            chat.box.on_set["AXSelectedText"] = self._typing_hook(chat)

    def _backspace(self, chat: FakeChat) -> None:
        current = self.box_text(chat.thread)
        location, length = self._selection(chat, current)
        start = location if length else max(location - 1, 0)
        updated = current[:start] + current[location + length :]
        chat.box.attributes["AXValue"] = updated or PLACEHOLDER + "\n"
        chat.box.attributes["AXSelectedTextRange"] = (start, 0)

    def _typing_hook(self, chat: FakeChat) -> Callable[[object], None]:
        def insert(value: object) -> None:
            self._insert(chat, str(value))

        return insert

    def _insert(self, chat: FakeChat, text: str) -> None:
        for old, new in self.options.substitutions.items():
            text = text.replace(old, new)
        current = self.box_text(chat.thread)
        location, length = self._selection(chat, current)
        updated = current[:location] + text + current[location + length :]
        chat.box.attributes["AXValue"] = updated or PLACEHOLDER + "\n"
        chat.box.attributes["AXSelectedTextRange"] = (location + len(text), 0)

    def _selection(self, chat: FakeChat, current: str) -> tuple[int, int]:
        selection = chat.box.attributes.get("AXSelectedTextRange")
        if isinstance(selection, tuple) and len(selection) == 2:
            location, length = int(selection[0]), int(selection[1])
            if 0 <= location <= len(current) and location + length <= len(current):
                return location, length
        return len(current), 0

    def _paste(self, chat: FakeChat, text: str) -> None:
        if self.options.send_on_newline and "\n" in text:
            first, rest = text.split("\n", 1)
            chat.messages.append(own(self._next_id(), first))
            self._insert(chat, rest)
            return
        self._insert(chat, text)

    def _commit(self) -> None:
        self.commits += 1
        chat = self.opened()
        if chat is None or not self.options.send_works:
            return
        typed = self.box_text(chat.thread)
        if chat.quote_text and typed.startswith(chat.quote_text):
            typed = typed[len(chat.quote_text) :]
        text = self.options.posted_text or typed.strip()
        if chat.quote is not None:
            text = (
                f"Begin quote, {chat.quote.author}, 10/20/2026, 9:00 AM, {chat.quote.text} {text}"
            )
        chat.messages.append(own(self._next_id(), text))
        chat.quote = None
        chat.quote_text = ""
        self.set_box(chat, "")

    def _next_id(self) -> int:
        ids = [
            int(item.id)
            for chat in self.chats
            for item in [*chat.messages, *chat.older]
            if item.id.isdigit()
        ]
        return max([FIRST_ID, *ids]) + 1000

    def _item(self, message_id: str, title: str) -> None:
        chat = self.opened()
        found = self._message(message_id)
        if chat is None or found is None:
            return
        if title == "Reply with quote" and self.options.quote_works:
            chat.quote = found
            if self.options.quote_in_value:
                chat.quote_text = f"{found.author}: {found.text}\n"
                self.set_box(chat, chat.quote_text)
        elif title in DELETE_TITLES:
            if self.options.delete_confirm:
                self.dialog = message_id
            elif self.options.delete_works:
                found.deleted = True
        elif title == "Mark as unread" and self.options.unread_works:
            chat.unread = True

    def _delete_confirmed(self) -> None:
        if self.dialog is None:
            return
        found = self._message(self.dialog)
        self.dialog = None
        if found is not None and self.options.delete_works:
            chat = self.opened()
            if chat is not None:
                chat.messages.remove(found)

    def _toggle(self, message_id: str, name: str) -> None:
        found = self._message(message_id)
        if found is None or not self.options.reaction_flips:
            return
        if name in found.reactions:
            found.reactions.discard(name)
        else:
            found.reactions.add(name)

    def _message(self, message_id: str) -> FakeMessage | None:
        chat = self.opened()
        if chat is None:
            return None
        return next((item for item in chat.messages if item.id == message_id), None)

    def _rows(self) -> list[Node]:
        return [
            chat_row(
                chat.thread,
                chat.name,
                chat.messages[-1].shown_text() if chat.messages else "",
                TIME_LABEL,
                kind=chat.kind,
                unread=chat.unread,
            )
            for chat in self.chats
            if chat.listed
        ]

    def _pane(self, chat: FakeChat) -> list[Node]:
        nodes: list[Node] = []
        for item in chat.messages:
            nodes.extend(message(item.id, item.author, item.shown_text(), STAMP, mine=item.mine))
            if self.options.popover and self.focused_message == item.id:
                nodes.append(self._popover(item))
        return nodes

    def _popover(self, item: FakeMessage) -> Node:
        toolbar = Node(
            role="AXToolbar",
            title="Reactions and message actions",
            children=tuple(self._toggle_node(item, name) for name in REACTIONS),
        )
        return Node(role="AXGroup", dom_id=f"{item.id}-popover-surface", children=(toolbar,))

    def _toggle_node(self, item: FakeMessage, name: str) -> Node:
        return Node(
            role="AXCheckBox",
            subrole="AXToggleButton",
            title=name,
            value=self._toggle_value(item, name),
            handle=("toggle", item.id, name),
        )

    def _toggle_value(self, item: FakeMessage, name: str) -> str:
        if not self.options.reaction_state_visible:
            return ""
        return "1" if name in item.reactions else "0"

    def _compose(self, chat: FakeChat) -> Node:
        box = Node(
            role="AXTextArea",
            title=PLACEHOLDER,
            value=str(chat.box.attributes.get("AXValue", "")),
            dom_id=f"new-message-{len(chat.name)}",
            handle=chat.box,
        )
        parts: list[Node] = [box]
        if chat.quote is not None and not self.options.quote_in_value:
            parts.append(
                Node(role="AXGroup", title=f"Begin quote, {chat.quote.author}, {chat.quote.text}")
            )
        if self.options.send_button:
            send = Node(role="AXButton", title="Send", handle=SEND_HANDLE)
            parts.append(Node(role="AXToolbar", title="Actions for new message", children=(send,)))
        return Node(role="AXGroup", children=tuple(parts))

    def _dialog(self) -> Node:
        buttons = (
            Node(role="AXButton", title="Cancel", handle=CANCEL_HANDLE),
            Node(role="AXButton", title="Delete", handle=CONFIRM_HANDLE),
        )
        return Node(
            role="AXGroup",
            subrole="AXApplicationDialog",
            title="Delete this message?",
            children=buttons,
        )
