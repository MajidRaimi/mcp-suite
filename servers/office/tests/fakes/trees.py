from __future__ import annotations

from dataclasses import dataclass

from fakes.ax_backend import FakeBackend, FakeElement, FakeProcesses, element
from fakes.clock import FakeClock, fake_waiter
from office_mcp.shared.ax import AccessMode, AxApp
from office_mcp.shared.node import Node

TEST_BUNDLE = "com.example.app"
TEST_APP_NAME = "Example App"
TEST_PID = 4242
LOAD_MORE = "Load more conversations"
CHAT_APP_ID = "86fcd49b-61a2-4701-b771-54728cd291fb"
ACTIVITY_APP_ID = "14d6962d-6eeb-4f48-8890-de55454bb136"


@dataclass
class AppRig:
    app: AxApp
    backend: FakeBackend
    processes: FakeProcesses
    clock: FakeClock
    root: FakeElement


def app_rig(
    *windows: FakeElement,
    mode: AccessMode = AccessMode.NATIVE,
    running: bool = True,
    menu_bar: FakeElement | None = None,
) -> AppRig:
    clock = FakeClock()
    backend = FakeBackend()
    root = element("AXApplication", *windows)
    if menu_bar is not None:
        root.add(menu_bar)
        root.attributes["AXMenuBar"] = menu_bar
    root.attributes["AXWindows"] = list(windows)
    backend.register(TEST_PID, root)
    processes = FakeProcesses({TEST_BUNDLE: TEST_PID} if running else {})
    app = AxApp(backend, processes, TEST_BUNDLE, TEST_APP_NAME, fake_waiter(clock), mode)
    return AppRig(app, backend, processes, clock, root)


def window(title: str, *kids: FakeElement, **attributes: object) -> FakeElement:
    return element("AXWindow", *kids, AXTitle=title, **attributes)


def button(label: str, **attributes: object) -> FakeElement:
    return element("AXButton", AXDescription=label, **attributes)


def menu_item(
    title: str,
    *submenu_items: FakeElement,
    enabled: bool = True,
    char: str = "",
    virtual_key: int | None = None,
    modifiers: int = 0,
    mark: str = "",
) -> FakeElement:
    attributes: dict[str, object] = {
        "AXTitle": title,
        "AXEnabled": enabled,
        "AXMenuItemCmdChar": char,
        "AXMenuItemCmdModifiers": modifiers,
        "AXMenuItemMarkChar": mark,
    }
    if virtual_key is not None:
        attributes["AXMenuItemCmdVirtualKey"] = virtual_key
    item = element("AXMenuItem", **attributes)
    if submenu_items:
        item.add(element("AXMenu", *submenu_items))
    return item


def menu_bar(*menus: tuple[str, list[FakeElement]]) -> FakeElement:
    bar = element("AXMenuBar")
    for title, items in menus:
        bar.add(element("AXMenuBarItem", element("AXMenu", *items), AXTitle=title))
    return bar


def popup_menu(*titles: str) -> FakeElement:
    return element("AXMenu", *(menu_item(title) for title in titles))


def outlook_row(description: str, handle: object | None = None) -> Node:
    return Node(role="AXRow", description=description, handle=handle or description)


def outlook_rows(*descriptions: str) -> list[Node]:
    return [outlook_row(text, handle=f"row{index}") for index, text in enumerate(descriptions)]


def load_more_row() -> Node:
    return Node(role="AXRow", description=LOAD_MORE, handle="load-more")


def folder_row(description: str, level: int = 1, handle: object | None = None) -> Node:
    return Node(role="AXRow", description=description, level=level, handle=handle or description)


def static(value: str) -> Node:
    return Node(role="AXStaticText", value=value)


def chat_row(
    thread: str,
    name: str,
    preview: str,
    time_label: str,
    kind: str = "Chat",
    presence: str = "Available",
    unread: bool = False,
    muted: bool = False,
) -> Node:
    parts = ["Unread message"] if unread else []
    parts += [kind, name]
    if kind == "Chat" and presence:
        parts.append(presence)
    parts += ["Last message", preview]
    if muted:
        parts.append("Muted")
    parts.append(time_label)
    children = [
        Node(role="AXGroup", dom_id=f"title-chat-list-item_{thread}", children=(static(name),)),
        Node(
            role="AXGroup", dom_id=f"time-chat-list-item_{thread}", children=(static(time_label),)
        ),
        Node(
            role="AXGroup",
            dom_id=f"message-preview-chat-list-item_{thread}",
            children=(static(preview),),
        ),
    ]
    if unread:
        children.insert(
            0, Node(role="AXGroup", dom_id="unread-label", children=(static("Unread"),))
        )
    return Node(role="AXRow", title=" ".join(parts), children=tuple(children), handle=thread)


def section_row(title: str, rows: list[Node] | None = None) -> Node:
    return Node(role="AXRow", title=title, children=tuple(rows or []), handle=title)


def message(
    message_id: str,
    author: str,
    body: str,
    stamp: str,
    mine: bool = False,
    show_author: bool = True,
) -> list[Node]:
    prefix = "fui-ChatMyMessage" if mine else "fui-ChatMessage"
    header: list[Node] = [
        Node(
            role="AXGroup",
            description=f"{stamp}.",
            dom_id=f"timestamp-{message_id}",
            classes=(f"{prefix}__timestamp",),
            children=(static("short"),),
        )
    ]
    if show_author:
        header.append(
            Node(role="AXGroup", dom_id=f"author-{message_id}", children=(static(author),))
        )
    return [
        Node(role="AXHeading", title=f"{body} by {author}", value="4"),
        Node(role="AXGroup", children=tuple(header)),
        Node(
            role="AXGroup",
            title=f"{body} {author} {stamp}.",
            dom_id=f"message-body-{message_id}",
            classes=(f"{prefix}__body",),
            children=(
                Node(role="AXButton", title="More message options"),
                Node(
                    role="AXGroup",
                    title=body,
                    dom_id=f"content-{message_id}",
                    children=(static(body),),
                ),
            ),
        ),
    ]


def rail(selected: str | None = CHAT_APP_ID) -> Node:
    toggles = tuple(
        Node(
            role="AXCheckBox",
            subrole="AXToggleButton",
            title=title,
            dom_id=app_id,
            value="1" if app_id == selected else "0",
            handle=f"rail:{title}",
        )
        for title, app_id in (("Activity", ACTIVITY_APP_ID), ("Chat", CHAT_APP_ID))
    )
    return Node(role="AXGroup", subrole="AXLandmarkNavigation", title="Apps", children=toggles)


def chat_header(thread: str, participants: list[str] | None = None) -> Node:
    people = tuple(
        Node(role="AXGroup", title=f"{name} Available", dom_id=f"chat-topic-person-8:orgid:{name}")
        for name in participants or []
    )
    roster = Node(role="AXList", title="Chat participants", children=people)
    return Node(role="AXGroup", dom_id=f"chat-header-{thread}", children=(roster,))


def teams_window(
    title: str,
    rows: list[Node],
    messages: list[Node],
    thread: str | None = None,
    selected_app: str | None = CHAT_APP_ID,
    extra: tuple[Node, ...] = (),
) -> Node:
    pane = Node(role="AXGroup", dom_id="chat-pane-list", children=tuple(messages))
    header = (chat_header(thread),) if thread is not None else ()
    return Node(
        role="AXWindow",
        title=title,
        children=(
            rail(selected_app),
            Node(role="AXGroup", children=tuple(rows)),
            *header,
            pane,
            *extra,
        ),
    )


def teams_menu(*titles: str) -> Node:
    items = tuple(Node(role="AXMenuItem", title=title, handle=f"item:{title}") for title in titles)
    return Node(role="AXMenu", children=items, handle="menu")


def calendar_button(title: str, handle: object | None = None) -> Node:
    return Node(role="AXButton", title=title, handle=handle or title)


def module_switcher() -> FakeElement:
    return element(
        "AXGroup",
        element("AXButton", AXHelp="Mail", AXSelected=True),
        element("AXButton", AXHelp="Calendar", AXSelected=False),
    )


def outlook_main_window(*kids: FakeElement, title: str = "Inbox • jane@example.com") -> FakeElement:
    return window(title, module_switcher(), *kids, AXSubrole="AXStandardWindow", AXMain=True)
