from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass

from office_mcp.features.outlook.actions.gateway import ActionsGateway, MoveSheet, MoveTarget
from office_mcp.features.outlook.models import Folder
from office_mcp.features.outlook.window import FocusedTab
from office_mcp.shared.errors import (
    AmbiguousTargetError,
    OfficeError,
    TargetError,
    TargetNotFoundError,
    UiTimeoutError,
)
from office_mcp.shared.events import Key, Modifier, Shortcut
from office_mcp.shared.scope import UiScope
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

MOVE_MENU = ("Message", "Move")
SHEET_ITEMS = frozenset({"choose folder...", "copy to folder..."})
LOCAL_ACCOUNT = "on my computer"
SHEET_CLOSE_WAIT = 5.0
MAX_LISTED = 12
CANCEL_SHEET = "cancel the Move sheet"
TAB_PATHS: dict[FocusedTab, tuple[tuple[str, ...], Shortcut]] = {
    FocusedTab.FOCUSED: (("Message", "Move to Focused"), Shortcut(Key.F, Modifier.OPTION)),
    FocusedTab.OTHER: (("Message", "Move to Other"), Shortcut(Key.O, Modifier.OPTION)),
}


@dataclass(frozen=True)
class FolderDestination:
    name: str
    path: str
    unique_name: bool

    @property
    def label(self) -> str:
        return self.path or self.name


@dataclass(frozen=True)
class TabDestination:
    tab: FocusedTab

    @property
    def label(self) -> str:
        return self.tab.value


Destination = FolderDestination | TabDestination


def resolve_destination(raw: str, folders: Sequence[Folder]) -> Destination:
    wanted = normalize(raw)
    if not wanted:
        raise TargetError("name the destination folder, or focused or other")
    for tab in FocusedTab:
        if wanted == tab.value:
            return TabDestination(tab)
    real = [folder for folder in folders if not folder.favorite]
    chosen = pick_folder(raw, real)
    same_name = [folder for folder in real if normalize(folder.name) == normalize(chosen.name)]
    return FolderDestination(chosen.name, chosen.path, unique_name=len(same_name) == 1)


def pick_folder(raw: str, folders: Sequence[Folder]) -> Folder:
    wanted = normalize(raw)
    matches = [
        folder for folder in folders if wanted in (normalize(folder.name), normalize(folder.path))
    ]
    by_path = [folder for folder in matches if normalize(folder.path) == wanted]
    if len(by_path) == 1:
        return by_path[0]
    if len(matches) == 1:
        return matches[0]
    if matches:
        raise AmbiguousTargetError(
            f"{raw!r} matches several folders; pass the full path",
            [folder.path for folder in matches],
        )
    names = ", ".join(dict.fromkeys(folder.path for folder in folders))
    raise TargetNotFoundError(f"no folder named {raw!r}. Folders: {names}")


def same_folder(source: str, destination: FolderDestination) -> bool:
    current = normalize(source)
    return bool(current) and current in (normalize(destination.name), normalize(destination.path))


def sheet_matches(
    targets: Sequence[MoveTarget], folder: str, account: str | None
) -> list[MoveTarget]:
    wanted = normalize(folder)
    owner = normalize(account) if account else ""
    return [
        target
        for target in targets
        if normalize(target.folder) == wanted
        and normalize(target.account) != LOCAL_ACCOUNT
        and (not owner or normalize(target.account) == owner)
    ]


def recent_items(titles: Sequence[str], folder: str) -> list[str]:
    wanted = normalize(folder)
    return [
        title
        for title in titles
        if normalize(title) not in SHEET_ITEMS and normalize(title) == wanted
    ]


class Mover:
    def __init__(
        self, gateway: ActionsGateway, waiter: Waiter, account: Callable[[], str | None]
    ) -> None:
        self._gateway = gateway
        self._waiter = waiter
        self._account = account

    def to_tab(self, tab: FocusedTab) -> str:
        path, shortcut = TAB_PATHS[tab]
        self._gateway.press_menu(path, shortcut)
        return f"moved to {tab.value} with the Message menu"

    def to_folder(self, scope: UiScope, destination: FolderDestination) -> str:
        recent = recent_items(self._gateway.menu_titles(MOVE_MENU), destination.name)
        if destination.unique_name and len(recent) == 1:
            self._gateway.press_menu((*MOVE_MENU, recent[0]))
            return f"moved to {destination.label} with the recent folders in the Move menu"
        sheet = self._gateway.open_move_sheet()
        scope.cleanup.push(CANCEL_SHEET, lambda: self._cancel(sheet))
        self._pick(sheet, destination)
        try:
            self._gateway.confirm_move(sheet)
            self._waiter.require(
                lambda: not self._gateway.sheet_open(sheet),
                lambda: UiTimeoutError("the Move sheet did not close after Move was pressed"),
                SHEET_CLOSE_WAIT,
            )
        except OfficeError as error:
            self._abandon(sheet, error)
            raise
        return f"moved to {destination.label} with the Move sheet"

    def _pick(self, sheet: MoveSheet, destination: FolderDestination) -> None:
        targets = self._gateway.filter_move_sheet(sheet, destination.name)
        matches = sheet_matches(targets, destination.name, self._account())
        try:
            if len(matches) != 1:
                raise sheet_error(destination, matches, targets)
            require_first(matches[0], [target.description for target in targets])
            self._gateway.pick_move_target(sheet, matches[0])
        except OfficeError as error:
            self._abandon(sheet, error)
            raise

    def _abandon(self, sheet: MoveSheet, error: OfficeError) -> None:
        try:
            self._cancel(sheet)
        except OfficeError as failure:
            error.add_note(f"the Move sheet could not be cancelled: {failure}")

    def _cancel(self, sheet: MoveSheet) -> None:
        if self._gateway.sheet_open(sheet):
            self._gateway.cancel_move(sheet)


def require_first(target: MoveTarget, listed: Sequence[str]) -> None:
    if listed and listed[0] == target.description:
        return
    above = list(listed[: listed.index(target.description)]) if target.description in listed else []
    raise AmbiguousTargetError(
        f"the Move sheet lists other folders above {target.description!r} and only its first "
        "entry can be picked safely, so nothing was moved; move it by hand in Outlook",
        above[:MAX_LISTED],
    )


def sheet_error(
    destination: FolderDestination,
    matches: Sequence[MoveTarget],
    targets: Sequence[MoveTarget],
) -> TargetError:
    if matches:
        return AmbiguousTargetError(
            f"the Move sheet lists {len(matches)} folders named {destination.name!r}, so "
            "nothing was moved",
            [target.description for target in matches[:MAX_LISTED]],
        )
    shown = ", ".join(target.description for target in targets[:MAX_LISTED]) or "none"
    return TargetNotFoundError(
        f"the Move sheet does not list {destination.name!r}, so nothing was moved. Shown: {shown}"
    )
