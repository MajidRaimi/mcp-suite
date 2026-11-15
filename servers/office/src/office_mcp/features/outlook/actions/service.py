from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field, replace
from datetime import date

from office_mcp.features.outlook.actions.flags import Flagger, FlagState
from office_mcp.features.outlook.actions.gateway import ActionsGateway
from office_mcp.features.outlook.actions.move import (
    Destination,
    FolderDestination,
    Mover,
    TabDestination,
    resolve_destination,
    same_folder,
)
from office_mcp.features.outlook.actions.neighbors import UnreadKeeper
from office_mcp.features.outlook.actions.watch import RowWatch, Target, is_unread
from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.listing import parse_folders
from office_mcp.features.outlook.models import Folder
from office_mcp.features.outlook.rows import KEYS_ONLY, RowFinder, SearchLocator, is_key
from office_mcp.features.outlook.view import button_titled
from office_mcp.features.outlook.window import OutlookMode, folder_from_title
from office_mcp.shared.errors import (
    AccessibilityError,
    ActionNotVerifiedError,
    ActivationError,
    AppNotRunningError,
    BusyError,
    ForbiddenActionError,
    MenuStuckError,
    OfficeError,
    OutlookUnavailableError,
    SessionUnavailableError,
    TargetError,
    error_text,
)
from office_mcp.shared.events import Key, Modifier, Shortcut
from office_mcp.shared.node import Node
from office_mcp.shared.results import ActionResult, ItemResult
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import ScopeFactory, ScopeOptions, UiScope
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

MAX_REFS = 25
ARCHIVE = "Archive"
ARCHIVE_FOLDER = "Archive"
DELETED_ITEMS = "Deleted Items"
ARCHIVE_MENU = ("Message", "Archive")
ARCHIVE_SHORTCUT = Shortcut(Key.E, Modifier.CONTROL)
DELETE_MENU = ("Edit", "Delete")
DELETE_SHORTCUT = Shortcut(Key.DELETE, Modifier.COMMAND)
ROW_ARCHIVE_UNVERIFIED = (
    "the Archive row action was pressed but the row is still listed, so it was not pressed "
    "again; not verified, check Archive"
)
ALERT_WAIT = 5.0
CANCEL = "Cancel"
PERMANENT = "permanent"
PROTECTED_FOLDERS = frozenset({"deleted items", "recoverable items"})
PATH_SEPARATOR = "/"
FOREGROUND = ScopeOptions(foreground=True)
ABORTING: tuple[type[OfficeError], ...] = (
    AccessibilityError,
    AppNotRunningError,
    OutlookUnavailableError,
    ActivationError,
    SessionUnavailableError,
    BusyError,
    ForbiddenActionError,
    MenuStuckError,
)

ItemAction = Callable[[UiScope, Target], ItemResult]


@dataclass
class BatchLog:
    items: list[ItemResult] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)


def validated_keys(refs: Sequence[str]) -> tuple[str, ...]:
    keys = tuple(dict.fromkeys(ref.strip() for ref in refs))
    if not keys:
        raise TargetError("pass at least one key from outlook_conversations or outlook_search")
    if len(keys) > MAX_REFS:
        raise TargetError(f"pass at most {MAX_REFS} keys per call")
    if not all(is_key(key) for key in keys):
        raise TargetError(KEYS_ONLY)
    return keys


def is_protected_folder(folder: str, folders: Sequence[Folder]) -> bool:
    wanted = normalize(folder)
    paths = [folder, *(item.path for item in folders if wanted in folder_keys(item))]
    return any(
        normalize(part) in PROTECTED_FOLDERS
        for path in paths
        for part in path.split(PATH_SEPARATOR)
    )


def protected_error(folder: str) -> ForbiddenActionError:
    return ForbiddenActionError(
        f"{folder!r} is Deleted Items or Recoverable Items, where Delete is permanent; "
        "nothing was deleted"
    )


def folder_keys(folder: Folder) -> tuple[str, str]:
    return (normalize(folder.name), normalize(folder.path))


def sheet_text(sheet: Node) -> str:
    pieces = (part for node in sheet.walk() for part in (node.title, node.value, node.description))
    return " ".join(" ".join(piece.split()) for piece in pieces if piece.strip())


def done_note(items: Sequence[ItemResult]) -> str:
    done = [item.ref for item in items if item.ok]
    listed = ", ".join(done) if done else "none"
    return f"done before the stop: {listed}"


class ActionsService:
    def __init__(
        self,
        outlook: OutlookGateway,
        gateway: ActionsGateway,
        rows: RowFinder,
        scopes: ScopeFactory,
        waiter: Waiter,
        today: Callable[[], date],
    ) -> None:
        self._outlook = outlook
        self._gateway = gateway
        self._rows = rows
        self._scopes = scopes
        self._waiter = waiter
        self._watch = RowWatch(outlook, waiter, today)
        self._keeper = UnreadKeeper(outlook, gateway, self._watch)
        self._flagger = Flagger(outlook, gateway, self._watch, self._keeper)
        self._mover = Mover(gateway, waiter, outlook.account)

    def archive(self, refs: Sequence[str], folder: str | None, deadline: Deadline) -> ActionResult:
        return self._batch("archive", refs, folder, deadline, lambda _scope: self._archive)

    def delete(self, refs: Sequence[str], folder: str | None, deadline: Deadline) -> ActionResult:
        result = self._batch("delete", refs, folder, deadline, lambda _scope: self._deleter(folder))
        return replace(result, extra={"moved_to": DELETED_ITEMS})

    def move(
        self, refs: Sequence[str], destination: str, folder: str | None, deadline: Deadline
    ) -> ActionResult:
        chosen: list[Destination] = []

        def prepare(_scope: UiScope) -> ItemAction:
            target = resolve_destination(destination, self._folders())
            chosen.append(target)
            return lambda scope, item: self._move(scope, item, target)

        result = self._batch("move", refs, folder, deadline, prepare)
        label = chosen[0].label if chosen else destination
        return replace(result, extra={"destination": label})

    def mark(
        self, refs: Sequence[str], read: bool, folder: str | None, deadline: Deadline
    ) -> ActionResult:
        action = "mark_read" if read else "mark_unread"
        return self._batch(
            action, refs, folder, deadline, lambda _scope: lambda _s, item: self._mark(item, read)
        )

    def flag(
        self, refs: Sequence[str], state: FlagState, folder: str | None, deadline: Deadline
    ) -> ActionResult:
        return self._batch(
            "flag",
            refs,
            folder,
            deadline,
            lambda _scope: lambda _s, item: self._flagger.apply(item, state),
        )

    def _batch(
        self,
        action: str,
        refs: Sequence[str],
        folder: str | None,
        deadline: Deadline,
        prepare: Callable[[UiScope], ItemAction],
    ) -> ActionResult:
        keys = validated_keys(refs)
        log = BatchLog()
        with self._scopes(FOREGROUND) as scope:
            handle = prepare(scope)
            for position, key in enumerate(keys):
                if deadline.expired():
                    log.skipped.extend(keys[position:])
                    break
                log.items.append(self._run_item(scope, key, folder, handle, log))
        return ActionResult(
            action=action,
            items=tuple(log.items),
            skipped=tuple(log.skipped),
            focus=scope.focus,
        )

    def _run_item(
        self,
        scope: UiScope,
        key: str,
        folder: str | None,
        handle: ItemAction,
        log: BatchLog,
    ) -> ItemResult:
        try:
            scope.ensure_front()
            locator = self._rows.locator_for(key, folder)
            row = self._rows.refresh(key, locator)
            result = handle(scope, Target(key, locator, row))
        except ABORTING as error:
            error.add_note(done_note(log.items))
            raise
        except OfficeError as error:
            return failed(key, error)
        return result

    def _archive(self, scope: UiScope, target: Target) -> ItemResult:
        if ARCHIVE not in self._outlook.row_actions(target.row.node):
            return self._relocate(target, ARCHIVE_FOLDER, self._menu_archive)
        self._outlook.perform_row_action(target.row.node, ARCHIVE)
        if self._watch.wait_gone(target.key):
            return moved(target, ARCHIVE_FOLDER, "archived with the row action")
        return moved(target, ARCHIVE_FOLDER, ROW_ARCHIVE_UNVERIFIED, ok=False)

    def _menu_archive(self) -> str:
        self._gateway.press_menu(ARCHIVE_MENU, ARCHIVE_SHORTCUT)
        return "archived with Message > Archive"

    def _deleter(self, hint: str | None) -> ItemAction:
        folders = self._folders()
        if hint and is_protected_folder(hint, folders):
            raise protected_error(hint)
        return lambda _s, target: self._delete(target, folders)

    def _delete(self, target: Target, folders: Sequence[Folder]) -> ItemResult:
        self._refuse_protected(target, folders)
        return self._relocate(target, DELETED_ITEMS, lambda: self._press_delete(target))

    def _refuse_protected(self, target: Target, folders: Sequence[Folder]) -> None:
        folder = target.folder
        if not folder and isinstance(target.locator, SearchLocator):
            raise ForbiddenActionError(
                "the search result does not say which folder it is in, so it was not deleted"
            )
        shown = ""
        if self._outlook.mode() is OutlookMode.MAIL:
            shown = folder_from_title(self._outlook.window_title())
        if any(is_protected_folder(name, folders) for name in (folder, shown) if name):
            raise protected_error(folder or shown)

    def _press_delete(self, target: Target) -> str:
        self._gateway.press_menu(DELETE_MENU, DELETE_SHORTCUT)
        self._waiter.until(
            lambda: bool(self._outlook.alert_sheets()) or self._watch.find(target.key) is None,
            ALERT_WAIT,
        )
        sheets = self._outlook.alert_sheets()
        if sheets:
            self._refuse_alert(sheets[0])
        count = target.row.summary.message_count
        noun = "message" if count == 1 else "messages"
        return f"{count} {noun} moved to {DELETED_ITEMS}"

    def _refuse_alert(self, sheet: Node) -> None:
        text = sheet_text(sheet)
        button = button_titled(sheet, CANCEL)
        if button is not None:
            self._outlook.press_button(button)
        state = "cancelled" if button is not None else "left open"
        if PERMANENT in text.casefold():
            raise ForbiddenActionError(
                f"Outlook asked to delete permanently and the alert was {state}; nothing was "
                f"deleted. Alert: {text[:200]!r}"
            )
        raise ActionNotVerifiedError(
            f"Outlook showed an alert after Delete and it was {state}: {text[:200]!r}"
        )

    def _move(self, scope: UiScope, target: Target, destination: Destination) -> ItemResult:
        if isinstance(destination, TabDestination):
            if target.row.summary.view == destination.tab.value:
                raise TargetError(f"it is already in the {destination.tab.value} tab")
            if self._outlook.focused_tab() is None:
                raise TargetError("this folder has no Focused and Other tabs")
            tab = destination.tab
            return self._relocate(target, destination.label, lambda: self._mover.to_tab(tab))
        return self._move_to_folder(scope, target, destination)

    def _move_to_folder(
        self, scope: UiScope, target: Target, destination: FolderDestination
    ) -> ItemResult:
        if same_folder(target.folder, destination):
            raise TargetError(f"it is already in {destination.label!r}")
        return self._relocate(
            target, destination.label, lambda: self._mover.to_folder(scope, destination)
        )

    def _relocate(self, target: Target, location: str, act: Callable[[], str]) -> ItemResult:
        watched = self._keeper.neighbors(target.key)
        notes: list[str] = []
        if not target.row.summary.unread:
            self._keeper.select(target.row)
        elif not self._keeper.mark_unread(target.row):
            notes.append("it could not be kept unread")
        self._keeper.require_selected(target.row)
        done = act()
        gone = self._watch.wait_gone(target.key)
        notes.extend(self._keeper.restore_neighbors(watched))
        if not gone:
            detail = "; ".join([f"still listed in {target.folder or 'the list'}", *notes])
            return moved(target, location, detail, ok=False)
        return moved(target, location, "; ".join([done, *notes]))

    def _mark(self, target: Target, read: bool) -> ItemResult:
        summary = target.row.summary
        wanted = "read" if read else "unread"
        if summary.unread != read:
            return stayed(target, True, f"already {wanted}")
        verified = (
            self._keeper.mark_read(target.row) if read else self._keeper.mark_unread(target.row)
        )
        self._outlook.deselect(target.row.node)
        if verified:
            return stayed(target, True, f"marked {wanted}")
        current = self._watch.find(target.key)
        state = (
            "unknown" if current is None else ("unread" if is_unread(current.summary) else "read")
        )
        return stayed(target, False, f"Outlook still shows it as {state}", ok=False)

    def _folders(self) -> list[Folder]:
        return parse_folders(self._outlook.folder_rows())


def failed(key: str, error: OfficeError) -> ItemResult:
    return ItemResult(
        ref=key, ok=False, label="", location="", verified=False, detail=error_text(error)
    )


def moved(target: Target, location: str, detail: str, ok: bool = True) -> ItemResult:
    return ItemResult(
        ref=target.key,
        ok=ok,
        label=target.label,
        location=location,
        verified=ok,
        detail=detail,
    )


def stayed(target: Target, verified: bool, detail: str, ok: bool = True) -> ItemResult:
    return ItemResult(
        ref=target.key,
        ok=ok,
        label=target.label,
        location=target.folder,
        verified=verified,
        detail=detail,
    )
