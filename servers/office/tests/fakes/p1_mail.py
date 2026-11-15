from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field, replace
from pathlib import Path

from fakes.drivers import FAKE_NOW, FAKE_TODAY, OutlookRig, outlook_rig
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.workspace import FakeWorkspace
from office_mcp.features.outlook.mail.search import SearchService
from office_mcp.features.outlook.mail.service import MailService
from office_mcp.features.outlook.people.service import PeopleService
from office_mcp.features.outlook.rows import SearchScope
from office_mcp.features.outlook.service import OutlookService
from office_mcp.features.outlook.window import FocusedTab
from office_mcp.shared.errors import UiTimeoutError
from office_mcp.shared.node import Node

UNREAD_PREFIX = "Unread, "
OWNER_APP = 100
OWNER_EMAIL = "jane@example.com"


def mail_row(
    sender: str,
    subject: str,
    time_label: str,
    preview: str,
    unread: bool = False,
    folder: str = "",
    flags: str = "",
) -> str:
    head = f"{UNREAD_PREFIX if unread else ''}{flags}  {sender}, {subject},"
    location = f"    Folder: {folder}," if folder else ""
    return f"{head}     {time_label},{location}        {preview}"


def section(title: str) -> str:
    return f"{title}, Expanded"


def static(value: str, description: str = "") -> Node:
    return Node(role="AXStaticText", value=value, description=description)


def attachment_group(name: str, size: str, menu: object | None = None) -> Node:
    parts = [static(name, "Attachment name"), static(size), Node(role="AXImage")]
    if menu is not None:
        parts.append(Node(role="AXPopUpButton", description="More Options", handle=menu))
    return Node(role="AXGroup", children=tuple(parts))


def attachment_list(*groups: Node) -> Node:
    label = f"{len(groups)} Attachment" + ("s" if len(groups) != 1 else "")
    return Node(role="AXList", description=label, children=(Node(role="AXList", children=groups),))


@dataclass(frozen=True)
class PaneSpec:
    subject: str
    sender: tuple[str, str] = ("Alice Example", "alice@example.com")
    to: tuple[tuple[str, str], ...] = (("Jane Doe", OWNER_EMAIL),)
    cc: tuple[tuple[str, str], ...] = ()
    sent: str = "Today at 9:15"
    body: tuple[str, ...] = ("Hello Jane,", "Numbers attached.")
    attachments: tuple[Node, ...] = ()


def recipient(name: str, email: str) -> Node:
    return Node(role="AXButton", value=f"{name}, {email}, Available")


def reading_pane(spec: PaneSpec) -> Node:
    header_children = [
        Node(role="AXGroup", value=f"From: {spec.sender[0]}, {spec.sender[1]}, Available"),
        static(spec.sent),
        static("To:"),
        *(recipient(name, email) for name, email in spec.to),
    ]
    if spec.cc:
        header_children.append(static("Cc:"))
        header_children.extend(recipient(name, email) for name, email in spec.cc)
    header = Node(
        role="AXScrollArea", description="Message Header Details", children=tuple(header_children)
    )
    top_children = [static(spec.subject), header]
    if spec.attachments:
        top_children.append(attachment_list(*spec.attachments))
    web = Node(
        role="AXWebArea",
        description="Reading Pane",
        children=tuple(static(line) for line in spec.body),
    )
    return Node(role="AXGroup", children=(Node(role="AXGroup", children=tuple(top_children)), web))


def read_version(node: Node) -> Node:
    if not node.description.startswith(UNREAD_PREFIX):
        return node
    return replace(node, description=node.description.removeprefix(UNREAD_PREFIX))


def unread_version(node: Node) -> Node:
    if node.description.startswith(UNREAD_PREFIX):
        return node
    return replace(node, description=UNREAD_PREFIX + node.description)


class MailGateway(FakeOutlookGateway):
    def __init__(
        self,
        rows: list[Node] | None = None,
        folders: list[Node] | None = None,
        tab: FocusedTab | None = FocusedTab.FOCUSED,
    ) -> None:
        super().__init__(rows=rows, folders=folders, tab=tab)
        self.on_press: Callable[[Node], None] | None = None
        self.read_on_deselect = False
        self.kept_unread: set[object] = set()
        self.expand_on_select: dict[object, str] = {}
        self.collapsed: dict[object, str] = {}
        self.collapse_ignored = False

    def select(self, node: Node) -> None:
        super().select(node)
        if not self.read_on_deselect:
            self._mark_read(node.handle)
        expanded = self.expand_on_select.get(node.handle)
        if expanded is not None:
            self._swap(node.handle, expanded)

    def collapse_selected(self) -> None:
        super().collapse_selected()
        if self.collapse_ignored:
            return
        for handle in self.selected:
            original = self.collapsed.pop(handle, None)
            if original is not None:
                self.rows = [
                    replace(row, description=original) if row.handle == handle else row
                    for row in self.rows
                ]

    def _swap(self, handle: object, description: str) -> None:
        for row in self.rows:
            if row.handle == handle and handle not in self.collapsed:
                self.collapsed[handle] = row.description
        self.rows = [
            replace(row, description=description) if row.handle == handle else row
            for row in self.rows
        ]

    def deselect(self, node: Node) -> None:
        super().deselect(node)
        if self.read_on_deselect and node.handle not in self.kept_unread:
            self._mark_read(node.handle)

    def _mark_read(self, handle: object) -> None:
        self.rows = [read_version(row) if row.handle == handle else row for row in self.rows]

    def press(self, node: Node) -> None:
        super().press(node)
        if self.on_press is not None:
            self.on_press(node)

    def row_handle(self, text: str) -> object:
        return next(row.handle for row in self.rows if text in row.description)

    def is_unread(self, handle: object) -> bool:
        return any(
            row.handle == handle and row.description.startswith(UNREAD_PREFIX) for row in self.rows
        )


class FakeUnreadMarker:
    def __init__(self, gateway: MailGateway, ignored: int = 0) -> None:
        self.gateway = gateway
        self.ignored = ignored
        self.presses = 0

    def mark_unread(self) -> None:
        self.presses += 1
        self.gateway.calls.append("mark_unread")
        if self.ignored > 0:
            self.ignored -= 1
            return
        selected = self.gateway.selected
        self.gateway.kept_unread |= selected
        self.gateway.rows = [
            unread_version(row) if row.handle in selected else row for row in self.gateway.rows
        ]


class FakeScopeControl:
    def __init__(self, gateway: FakeOutlookGateway, scope: SearchScope | None) -> None:
        self.gateway = gateway
        self.scope = scope
        self.chosen: list[SearchScope] = []

    def current(self) -> SearchScope | None:
        if self.gateway.search is not None:
            return self.gateway.search.scope
        return self.scope

    def choose(self, scope: SearchScope) -> None:
        self.chosen.append(scope)
        self.gateway.calls.append(f"choose_scope:{scope.value}")
        self.scope = scope
        if self.gateway.search is not None:
            self.gateway.search = replace(self.gateway.search, scope=scope)


class FakeSuggestions:
    def __init__(self, rows: list[str] | None = None, fail: bool = False) -> None:
        self.typed: list[str] = []
        self.shown: list[str] = []
        self.results = list(rows or [])
        self.fail = fail
        self.cleared = 0

    def type(self, query: str) -> None:
        if self.fail:
            raise UiTimeoutError("the search field did not take the text")
        self.typed.append(query)
        self.shown = list(self.results)

    def rows(self) -> list[str]:
        return list(self.shown)

    def clear(self) -> None:
        self.cleared += 1
        self.shown = []


@dataclass
class FakeDownloads:
    directory: Path
    content: bytes = b"x" * 2048
    snapshots: int = 0
    waited: list[str] = field(default_factory=list)

    def snapshot(self) -> frozenset[str]:
        self.snapshots += 1
        return frozenset(entry.name for entry in self.directory.iterdir())

    def wait_for(self, name: str, before: frozenset[str], timeout: float) -> Path:
        self.waited.append(name)
        path = self.directory / name
        path.write_bytes(self.content)
        return path


@dataclass
class MailRig:
    rig: OutlookRig
    gateway: MailGateway
    marker: FakeUnreadMarker
    core: OutlookService
    control: FakeScopeControl
    suggestions: FakeSuggestions
    mail: MailService


def mail_rig(
    gateway: MailGateway,
    scope: SearchScope | None = SearchScope.MAILBOX,
    suggestions: FakeSuggestions | None = None,
    ignored_unread: int = 0,
) -> MailRig:
    rig = outlook_rig(gateway, workspace=FakeWorkspace(frontmost=OWNER_APP))
    driver = rig.driver
    marker = FakeUnreadMarker(gateway, ignored_unread)
    core = OutlookService(
        gateway,
        driver.rows,
        driver.scope,
        driver.waiter,
        lambda: FAKE_TODAY,
        lambda: FAKE_NOW,
        marker,
    )
    control = FakeScopeControl(gateway, scope)
    source = suggestions or FakeSuggestions()
    mail = MailService(
        core,
        SearchService(core, driver.rows, gateway, control),
        PeopleService(source, driver.scope, driver.waiter),
        None,
        lambda: OWNER_EMAIL,
        lambda: FAKE_TODAY,
    )
    return MailRig(rig, gateway, marker, core, control, source, mail)
