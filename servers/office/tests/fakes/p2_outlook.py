from __future__ import annotations

from dataclasses import dataclass, field

from fakes.clock import FakeClock, fake_waiter
from fakes.drivers import FAKE_TODAY as TODAY
from fakes.outlook_gateway import FakeOutlookGateway
from fakes.scope import recording_factory
from fakes.trees import folder_row
from office_mcp.features.outlook.actions.gateway import MoveSheet, MoveTarget
from office_mcp.features.outlook.actions.service import ActionsService
from office_mcp.features.outlook.listing import parse_rows
from office_mcp.features.outlook.rows import RefRegistry, RowFinder
from office_mcp.features.outlook.window import FocusedTab
from office_mcp.shared.errors import ActionNotVerifiedError, ElementNotFoundError
from office_mcp.shared.events import Shortcut
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.node import Node
from office_mcp.shared.runner import Deadline

DELETE = ("Edit", "Delete")
ARCHIVE = ("Message", "Archive")
MARK_UNREAD = ("Message", "Mark as Unread")
MARK_READ = ("Message", "Mark as Read")
TAB_MOVES = (("Message", "Move to Focused"), ("Message", "Move to Other"))
SHEET_TITLES = ("Choose Folder...", "Copy to Folder...")
CLEAR_FLAG = "Clear Flag"


@dataclass
class Mail:
    sender: str
    subject: str
    time: str
    preview: str
    unread: bool = False
    flagged: bool = False
    count: int = 1
    folder: str = ""

    def description(self) -> str:
        flags = (["Unread"] if self.unread else []) + (["Flagged"] if self.flagged else [])
        counted = [f"{self.count} messages"] if self.count > 1 else []
        head = ", ".join([*flags, *counted, self.sender, self.subject])
        located = f" Folder: {self.folder}," if self.folder else ""
        return f"  {head},     {self.time},{located}        {self.preview}"


def atlas(unread: bool = False, flagged: bool = False, count: int = 1, folder: str = "") -> Mail:
    return Mail(
        "Alice Example",
        "Project Atlas budget",
        "9:15 AM",
        "Numbers attached",
        unread,
        flagged,
        count,
        folder,
    )


def offsite(unread: bool = False) -> Mail:
    return Mail("Bob Sample", "Team offsite", "8:40 AM", "Agenda inside", unread)


def lunch(unread: bool = False) -> Mail:
    return Mail("Jane Doe", "Lunch on Thursday", "Yesterday", "Works for me", unread)


class OrganizeOutlook(FakeOutlookGateway):
    def __init__(self, mails: dict[str, Mail], folders: list[Node] | None = None) -> None:
        super().__init__(folders=folders or default_folders())
        self.mails = dict(mails)
        self.select_marks_read = True
        self.row_archive_works = True
        self.auto_select_next = False
        self.select_ignored = False
        self.first_select_lands_on_child: set[str] = set()
        self.select_extends = False
        self.hidden_selected: set[str] = set()
        self.multi_pane: Node | None = None

    def message_rows(self) -> list[Node]:
        self.calls.append("message_rows")
        return [
            Node(
                role="AXRow",
                description=mail.description(),
                handle=handle,
                selected=handle in self.selected,
            )
            for handle, mail in self.mails.items()
        ]

    def select(self, node: Node) -> None:
        if self.select_ignored:
            self.calls.append(f"ignored_select:{node.handle}")
            return
        if str(node.handle) in self.first_select_lands_on_child:
            self.first_select_lands_on_child.discard(str(node.handle))
            self.calls.append(f"child_select:{node.handle}")
            return
        kept = set(self.selected) if self.select_extends else set()
        super().select(node)
        self.selected |= kept
        mail = self.mails.get(str(node.handle))
        if mail is not None and self.select_marks_read:
            mail.unread = False

    def reading_pane(self) -> Node | None:
        if self.multi_pane is not None and len(self.selected | self.hidden_selected) > 1:
            return self.multi_pane
        return super().reading_pane()

    def perform_row_action(self, row: Node, label: str) -> None:
        super().perform_row_action(row, label)
        if label == "Archive" and self.row_archive_works:
            self.remove(str(row.handle))

    def remove(self, handle: str | None) -> None:
        if handle is None or handle not in self.mails:
            return
        order = list(self.mails)
        position = order.index(handle)
        del self.mails[handle]
        self.selected.discard(handle)
        rest = list(self.mails)
        if self.auto_select_next and position < len(rest):
            following = rest[position]
            self.selected = {following}
            if self.select_marks_read:
                self.mails[following].unread = False

    def selected_handle(self) -> str | None:
        return str(next(iter(self.selected))) if self.selected else None

    def key_of(self, handle: str) -> str:
        node = Node(role="AXRow", description=self.mails[handle].description(), handle=handle)
        return parse_rows([node], TODAY)[0].summary.key


def default_folders() -> list[Node]:
    return [
        folder_row("Favorites", level=0),
        folder_row("Inbox; 3 unread messages", level=1, handle="fav-inbox"),
        folder_row("jane@example.com", level=0),
        folder_row("Inbox; 3 unread messages", level=1, handle="inbox"),
        folder_row("Archive", level=1),
        folder_row("Deleted Items; 2 unread messages", level=1),
        folder_row("Old", level=2),
        folder_row("Projects", level=1),
        folder_row("Clients", level=1),
        folder_row("Projects", level=2, handle="clients-projects"),
        folder_row("Travel", level=1),
    ]


@dataclass
class FakeActions:
    outlook: OrganizeOutlook
    presses: list[tuple[tuple[str, ...], Shortcut | None]] = field(default_factory=list)
    inert: set[tuple[str, ...]] = field(default_factory=set)
    move_titles: tuple[str, ...] = SHEET_TITLES
    targets: tuple[MoveTarget, ...] = ()
    pick_fails: bool = False
    hides_offered_items: bool = False
    move_works: bool = True
    sheet_sticks: bool = False
    flag_works: bool = True
    unread_sticks: bool = True
    flag_checks: dict[str, bool | None] = field(default_factory=dict)
    alert: Node | None = None
    deletes_before_alert: int = 0
    sheet: bool = False

    def press_menu(self, path: tuple[str, ...], shortcut: Shortcut | None = None) -> None:
        self.presses.append((path, shortcut))
        self.log("menu:" + " > ".join(path))
        if path in self.inert:
            return
        selected = self.outlook.selected_handle()
        if (
            self.hides_offered_items
            and path[:-1] == MARK_UNREAD[:-1]
            and path[-1] not in self._message_titles()
        ):
            raise ElementNotFoundError(f"menu item {' > '.join(path)} was not found")
        if path == DELETE:
            self._delete(selected)
        elif path in (ARCHIVE, *TAB_MOVES) or self._recent_move(path):
            self.outlook.remove(selected)
        elif path in (MARK_UNREAD, MARK_READ) and selected in self.outlook.mails:
            self._mark(str(selected), unread=path == MARK_UNREAD)

    def _mark(self, handle: str, unread: bool) -> None:
        if unread and not self.unread_sticks:
            return
        self.outlook.mails[handle].unread = unread

    def menu_titles(self, path: tuple[str, ...]) -> tuple[str, ...]:
        self.log("menu_titles:" + " > ".join(path))
        if path == MARK_UNREAD[:-1]:
            return self._message_titles()
        return self.move_titles

    def _message_titles(self) -> tuple[str, ...]:
        mail = self.outlook.mails.get(self.outlook.selected_handle() or "")
        if mail is None:
            return ()
        return (MARK_READ[-1],) if mail.unread else (MARK_UNREAD[-1],)

    def open_row_menu(self) -> Node:
        self.log("row_menu")
        return Node(role="AXMenu", handle="row-menu")

    def choose(self, menu: Node, path: tuple[str, ...]) -> None:
        self.log("choose:" + " > ".join(path))
        mail = self.outlook.mails.get(self.outlook.selected_handle() or "")
        if mail is None or not self.flag_works or path[0] != "Flag":
            return
        mail.flagged = not mail.flagged if path[-1] == CLEAR_FLAG else True

    def checked(self, menu: Node, path: tuple[str, ...]) -> bool | None:
        return self.flag_checks.get(path[-1])

    def close_menu(self, menu: Node | None = None) -> None:
        self.log("close_menu")

    def open_move_sheet(self) -> MoveSheet:
        self.log("open_sheet")
        self.sheet = True
        return MoveSheet("sheet")

    def filter_move_sheet(self, sheet: MoveSheet, text: str) -> tuple[MoveTarget, ...]:
        self.log(f"filter:{text}")
        return self.targets

    def pick_move_target(self, sheet: MoveSheet, target: MoveTarget) -> None:
        self.log(f"pick:{target.description}")
        if self.pick_fails:
            raise ActionNotVerifiedError(
                f"the Move sheet changed while picking {target.folder!r}; nothing was moved"
            )

    def confirm_move(self, sheet: MoveSheet) -> None:
        self.log("confirm_move")
        if self.sheet_sticks:
            return
        self.sheet = False
        if self.move_works:
            self.outlook.remove(self.outlook.selected_handle())

    def cancel_move(self, sheet: MoveSheet) -> None:
        self.log("cancel_sheet")
        self.sheet = False

    def sheet_open(self, sheet: MoveSheet) -> bool:
        return self.sheet

    def log(self, call: str) -> None:
        self.outlook.calls.append(call)

    def _delete(self, selected: str | None) -> None:
        if self.alert is not None and self.deletes_before_alert <= 0:
            self.outlook.sheets.append(self.alert)
            return
        self.deletes_before_alert -= 1
        self.outlook.remove(selected)

    def _recent_move(self, path: tuple[str, ...]) -> bool:
        return len(path) == 3 and path[:2] == ("Message", "Move") and path[2] not in SHEET_TITLES


@dataclass
class ActionsRig:
    service: ActionsService
    outlook: OrganizeOutlook
    actions: FakeActions
    scope_log: list[str]
    clock: FakeClock
    registry: RefRegistry

    def deadline(self, seconds: float = 60.0) -> Deadline:
        return Deadline.after(self.clock, seconds)

    def key(self, handle: str) -> str:
        return self.outlook.key_of(handle)

    def visible(self) -> list[str]:
        return [call for call in self.outlook.calls if call not in ("message_rows", "folder_rows")]


def actions_rig(
    mails: dict[str, Mail],
    folder: str = "Inbox",
    tab: FocusedTab | None = FocusedTab.FOCUSED,
    outcome: FocusOutcome = FocusOutcome.RESTORED,
) -> ActionsRig:
    outlook = OrganizeOutlook(mails)
    outlook.folder = folder
    outlook.tab = tab
    actions = FakeActions(outlook)
    clock = FakeClock()
    waiter = fake_waiter(clock)
    registry = RefRegistry()
    factory, log = recording_factory(outcome)
    service = ActionsService(
        outlook=outlook,
        gateway=actions,
        rows=RowFinder(outlook, waiter, lambda: TODAY, registry),
        scopes=factory,
        waiter=waiter,
        today=lambda: TODAY,
    )
    return ActionsRig(service, outlook, actions, log, clock, registry)


def deleted_items_alert() -> Node:
    return Node(
        role="AXSheet",
        children=(
            Node(
                role="AXStaticText",
                value="These items will be permanently deleted. Do you want to continue?",
            ),
            Node(role="AXButton", title="Cancel"),
            Node(role="AXButton", title="Delete"),
        ),
    )
