from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.listing import folder_pairs
from office_mcp.features.outlook.rows import SearchScope, SearchState, open_folder
from office_mcp.features.outlook.window import (
    FocusedTab,
    OutlookMode,
    OutlookModule,
    folder_from_title,
)
from office_mcp.shared.errors import CleanupError, OfficeError, error_text
from office_mcp.shared.node import Node
from office_mcp.shared.text import normalize
from office_mcp.shared.waiting import Waiter

SHEET_CANCEL_TITLES = ("Cancel", "Don't Send", "Continue Editing")
FOLDER_SCOPES = frozenset({SearchScope.CURRENT_FOLDER, SearchScope.SUBFOLDERS})


@dataclass(frozen=True)
class ViewState:
    mode: OutlookMode
    folder: str
    tab: FocusedTab | None
    search: SearchState | None
    had_selection: bool


def capture_view(gateway: OutlookGateway) -> ViewState:
    mode = gateway.mode()
    return ViewState(
        mode=mode,
        folder=shown_folder(gateway, mode),
        tab=gateway.focused_tab() if mode is not OutlookMode.CALENDAR else None,
        search=gateway.search_state() if mode is OutlookMode.SEARCH else None,
        had_selection=mode is not OutlookMode.CALENDAR and has_selection(gateway),
    )


def shown_folder(gateway: OutlookGateway, mode: OutlookMode) -> str:
    if mode is OutlookMode.MAIL:
        return folder_from_title(gateway.window_title())
    if mode is OutlookMode.SEARCH:
        return selected_folder_path(gateway)
    return ""


def selected_folder_path(gateway: OutlookGateway) -> str:
    return next(
        (folder.path for node, folder in folder_pairs(gateway.folder_rows()) if node.selected),
        "",
    )


def has_selection(gateway: OutlookGateway) -> bool:
    return any(node.selected for node in gateway.message_rows())


def restore_view(gateway: OutlookGateway, state: ViewState, waiter: Waiter) -> tuple[str, ...]:
    steps: tuple[Callable[[], str | None], ...] = (
        lambda: close_leftovers(gateway),
        lambda: restore_search(gateway, state, waiter),
        lambda: restore_module(gateway, state),
        lambda: restore_folder(gateway, state, waiter),
        lambda: restore_tab(gateway, state),
        lambda: restore_selection(gateway, state),
    )
    notes: list[str] = []
    failures: list[str] = []
    for step in steps:
        try:
            note = step()
        except OfficeError as error:
            failures.append(error_text(error))
            continue
        if note:
            notes.append(note)
    if failures:
        raise CleanupError("could not put the Outlook view back: " + "; ".join(failures))
    return tuple(notes)


def close_leftovers(gateway: OutlookGateway) -> str | None:
    dismissed = gateway.dismiss_known_popovers() + gateway.close_search_suggestions()
    cancelled = sum(cancel_sheet(gateway, sheet) for sheet in gateway.alert_sheets())
    if not dismissed and not cancelled:
        return None
    return f"closed {dismissed + cancelled} leftover popover or sheet"


def cancel_sheet(gateway: OutlookGateway, sheet: Node) -> int:
    for title in SHEET_CANCEL_TITLES:
        button = button_titled(sheet, title)
        if button is not None:
            gateway.press_button(button)
            return 1
    raise CleanupError(f"an Outlook sheet is still open: {sheet.label()[:120]!r}")


def button_titled(root: Node, title: str) -> Node | None:
    return root.find(lambda node: node.role == "AXButton" and node.label() == title)


def restore_search(gateway: OutlookGateway, state: ViewState, waiter: Waiter) -> str | None:
    current = gateway.search_state()
    if current == state.search:
        return None
    if state.search is None:
        gateway.cancel_search()
        return "left search"
    if state.folder and state.search.scope in FOLDER_SCOPES:
        if current is not None:
            gateway.cancel_search()
        open_folder(gateway, waiter, state.folder)
    gateway.run_search(state.search.query, state.search.scope)
    return "ran the previous search again"


def restore_module(gateway: OutlookGateway, state: ViewState) -> str | None:
    calendar_now = gateway.mode() is OutlookMode.CALENDAR
    calendar_before = state.mode is OutlookMode.CALENDAR
    if calendar_now == calendar_before:
        return None
    module = OutlookModule.CALENDAR if calendar_before else OutlookModule.MAIL
    gateway.show_module(module)
    return f"back to {module.value}"


def restore_folder(gateway: OutlookGateway, state: ViewState, waiter: Waiter) -> str | None:
    if state.mode is not OutlookMode.MAIL or not state.folder:
        return None
    if normalize(folder_from_title(gateway.window_title())) == normalize(state.folder):
        return None
    open_folder(gateway, waiter, state.folder)
    return f"reopened {state.folder}"


def restore_tab(gateway: OutlookGateway, state: ViewState) -> str | None:
    if state.tab is None or state.mode is OutlookMode.CALENDAR:
        return None
    if gateway.focused_tab() in (None, state.tab):
        return None
    gateway.switch_tab(state.tab)
    return f"back to the {state.tab.value} tab"


def restore_selection(gateway: OutlookGateway, state: ViewState) -> str | None:
    if state.had_selection or state.mode is OutlookMode.CALENDAR:
        return None
    selected = [node for node in gateway.message_rows() if node.selected]
    for node in selected:
        gateway.deselect(node)
    return "cleared the selection" if selected else None
