from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from enum import StrEnum
from typing import Any, Protocol

from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.models import Summary
from office_mcp.features.outlook.rows import FolderLocator, RowFinder, SearchLocator, SearchScope
from office_mcp.features.outlook.service import OutlookService
from office_mcp.features.outlook.window import OutlookMode, OutlookModule
from office_mcp.shared.ax import AxApp
from office_mcp.shared.errors import ElementNotFoundError, UiTimeoutError
from office_mcp.shared.menus import PopupMenus
from office_mcp.shared.runner import Deadline
from office_mcp.shared.scope import UiScope
from office_mcp.shared.tree import by_role_and_label, find
from office_mcp.shared.waiting import Waiter

SEARCH_FIELD = "Search"
SEARCH_SCOPE = "Search Scope"
FIELD_DEPTH = 4
FOCUS_SECONDS = 2.0
HEAVY_ROLES = frozenset({"AXTable", "AXOutline", "AXWebArea"})
LEAVE_SEARCH = "leave the search"
RESTORE_SCOPE = "put the search scope back"


class SearchArea(StrEnum):
    FOLDER = "folder"
    SUBFOLDERS = "subfolders"
    MAILBOX = "mailbox"
    ALL = "all"


AREA_SCOPES: dict[SearchArea, SearchScope] = {
    SearchArea.FOLDER: SearchScope.CURRENT_FOLDER,
    SearchArea.SUBFOLDERS: SearchScope.SUBFOLDERS,
    SearchArea.MAILBOX: SearchScope.MAILBOX,
    SearchArea.ALL: SearchScope.ALL,
}
FOLDER_AREAS = frozenset({SearchArea.FOLDER, SearchArea.SUBFOLDERS})


@dataclass(frozen=True)
class SearchResult:
    query: str
    area: SearchArea
    results: tuple[Summary, ...]

    def to_record(self) -> dict[str, Any]:
        return {
            "query": self.query,
            "scope": self.area.value,
            "count": len(self.results),
            "results": [summary.to_record() for summary in self.results],
        }


class SearchScopeControl(Protocol):
    def current(self) -> SearchScope | None: ...

    def choose(self, scope: SearchScope) -> None: ...


def scope_from(value: str) -> SearchScope | None:
    try:
        return SearchScope(value.strip())
    except ValueError:
        return None


def search_field(app: AxApp, window: object) -> object | None:
    return find(
        app,
        window,
        lambda element: (
            app.text(element, "AXRole") == "AXTextField"
            and app.text(element, "AXDescription") == SEARCH_FIELD
        ),
        FIELD_DEPTH,
        prune=lambda element: app.text(element, "AXRole") in HEAVY_ROLES,
    )


class AxSearchScope:
    def __init__(
        self, app: AxApp, window: Callable[[], object], popups: PopupMenus, waiter: Waiter
    ) -> None:
        self._app = app
        self._window = window
        self._popups = popups
        self._waiter = waiter

    def current(self) -> SearchScope | None:
        button = self._button()
        return scope_from(self._app.text(button, "AXValue")) if button is not None else None

    def choose(self, scope: SearchScope) -> None:
        button = self._button()
        if button is None:
            raise ElementNotFoundError("the Outlook search scope menu was not found")
        menu = self._popups.open(button, "AXPress")
        try:
            self._popups.choose(menu, (scope.value,))
        finally:
            self._popups.close()
        self._waiter.require(
            lambda: self.current() is scope,
            lambda: UiTimeoutError(f"Outlook did not switch the search scope to {scope.value}"),
        )

    def _button(self) -> object | None:
        field = search_field(self._app, self._window())
        if field is None:
            return None
        button = self._scope_button(field)
        if button is not None:
            return button
        self._app.set_attribute(field, "AXFocused", True)
        return self._waiter.until(lambda: self._scope_button(field), FOCUS_SECONDS)

    def _scope_button(self, field: object) -> object | None:
        return find(
            self._app,
            field,
            by_role_and_label(self._app, "AXPopUpButton", SEARCH_SCOPE),
            FIELD_DEPTH,
        )


class SearchService:
    def __init__(
        self,
        core: OutlookService,
        rows: RowFinder,
        gateway: OutlookGateway,
        control: SearchScopeControl,
    ) -> None:
        self._core = core
        self._rows = rows
        self._gateway = gateway
        self._control = control

    def search(
        self,
        query: str,
        area: SearchArea,
        folder: str | None,
        limit: int,
        load_more: int,
        deadline: Deadline | None = None,
    ) -> SearchResult:
        scope = AREA_SCOPES[area]
        with self._core.scope() as ui:
            self._show_mail(folder if area in FOLDER_AREAS else None)
            ui.cleanup.push(LEAVE_SEARCH, self._gateway.cancel_search)
            self._keep_scope(ui, scope)
            locator = SearchLocator(query, scope)
            summaries = self._core.list_rows(locator, None, load_more, deadline)
        return SearchResult(query, area, tuple(summaries[: max(limit, 0)]))

    def _show_mail(self, folder: str | None) -> None:
        if folder:
            self._rows.rows(FolderLocator(folder))
        elif self._gateway.mode() is OutlookMode.CALENDAR:
            self._gateway.show_module(OutlookModule.MAIL)

    def _keep_scope(self, ui: UiScope, wanted: SearchScope) -> None:
        previous = self._control.current()
        if previous is None or previous is wanted:
            return
        ui.cleanup.push(RESTORE_SCOPE, lambda: self._restore(previous))

    def _restore(self, previous: SearchScope) -> None:
        if self._control.current() is not previous:
            self._control.choose(previous)
