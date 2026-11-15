from __future__ import annotations

from dataclasses import replace

from fakes.trees import LOAD_MORE
from office_mcp.features.outlook.listing import parse_folder
from office_mcp.features.outlook.rows import SearchScope, SearchState
from office_mcp.features.outlook.window import FocusedTab, OutlookMode, OutlookModule
from office_mcp.shared.errors import ElementNotFoundError
from office_mcp.shared.node import Node
from office_mcp.shared.windows import WindowProbe

ACCOUNT = "jane@example.com"
HEALTHY_LABELS = ("New Mail; Press space for other actions",)


class FakeOutlookGateway:
    def __init__(
        self,
        rows: list[Node] | None = None,
        folders: list[Node] | None = None,
        folder: str = "Inbox",
        tab: FocusedTab | None = FocusedTab.FOCUSED,
        labels: tuple[str, ...] = HEALTHY_LABELS,
    ) -> None:
        self.rows: list[Node] = list(rows or [])
        self.folders: list[Node] = list(folders or [])
        self.folder = folder
        self.tab = tab
        self.labels = list(labels)
        self.mode_value = OutlookMode.MAIL
        self.search: SearchState | None = None
        self.folder_rows_by_name: dict[str, list[Node]] = {}
        self.tab_rows: dict[FocusedTab, list[Node]] = {}
        self.search_rows: list[Node] = []
        self.pages: list[list[Node]] = []
        self.panes: dict[object, Node] = {}
        self.actions: dict[object, tuple[str, ...]] = {}
        self.sheets: list[Node] = []
        self.popovers = 0
        self.suggestions = 0
        self.selected: set[object] = set()
        self.calls: list[str] = []
        self.launches: list[bool] = []
        self.fail_labels: str | None = None

    def main_probe(self) -> WindowProbe:
        return WindowProbe(
            handle="main",
            title=self.window_title(),
            subrole="AXStandardWindow",
            main=True,
            focused=True,
            landmarks=frozenset({"Mail", "Calendar", "Message List"}),
        )

    def window_title(self) -> str:
        if self.mode_value is OutlookMode.SEARCH:
            return f'Searching "{ACCOUNT}"'
        if self.mode_value is OutlookMode.CALENDAR:
            return "Calendar"
        return f"{self.folder} • {ACCOUNT}"

    def account(self) -> str | None:
        return ACCOUNT

    def mode(self) -> OutlookMode:
        return self.mode_value

    def shallow_button_labels(self, launch: bool = True) -> list[str]:
        self.launches.append(launch)
        if self.fail_labels is not None:
            raise ElementNotFoundError(self.fail_labels)
        return list(self.labels)

    def folder_rows(self) -> list[Node]:
        self.calls.append("folder_rows")
        return [self._marked(node) for node in self.folders]

    def message_rows(self) -> list[Node]:
        self.calls.append("message_rows")
        return [self._marked(node) for node in self.rows]

    def select(self, node: Node) -> None:
        self.calls.append(f"select:{node.handle}")
        if any(node.handle == folder.handle for folder in self.folders):
            self._open_folder(node)
            return
        self.selected = {node.handle}

    def deselect(self, node: Node) -> None:
        self.calls.append(f"deselect:{node.handle}")
        self.selected.discard(node.handle)

    def collapse_selected(self) -> None:
        self.calls.append("collapse")

    def press(self, node: Node) -> None:
        self.calls.append(f"press:{node.handle}")
        if node.description == LOAD_MORE and self.pages:
            self.rows = [row for row in self.rows if row.description != LOAD_MORE]
            self.rows.extend(self.pages.pop(0))

    def reading_pane(self) -> Node | None:
        handle = next(iter(self.selected), None)
        return self.panes.get(handle) if handle is not None else None

    def toolbar_button(self, name: str) -> Node | None:
        return None

    def press_button(self, node: Node) -> None:
        self.calls.append(f"press_button:{node.label()}")
        self.sheets = [sheet for sheet in self.sheets if sheet.find(lambda n: n is node) is None]

    def focused_tab(self) -> FocusedTab | None:
        return self.tab

    def switch_tab(self, tab: FocusedTab) -> None:
        self.calls.append(f"switch_tab:{tab.value}")
        self.tab = tab
        if tab in self.tab_rows:
            self.rows = list(self.tab_rows[tab])

    def search_state(self) -> SearchState | None:
        return self.search

    def run_search(self, query: str, scope: SearchScope) -> None:
        self.calls.append(f"run_search:{query}:{scope.value}")
        self.search = SearchState(query, scope)
        self.mode_value = OutlookMode.SEARCH
        self.rows = list(self.search_rows)

    def cancel_search(self) -> None:
        self.calls.append("cancel_search")
        self.search = None
        self.mode_value = OutlookMode.MAIL

    def show_module(self, module: OutlookModule) -> None:
        self.calls.append(f"show_module:{module.value}")
        self.mode_value = (
            OutlookMode.CALENDAR if module is OutlookModule.CALENDAR else OutlookMode.MAIL
        )

    def row_actions(self, row: Node) -> tuple[str, ...]:
        return self.actions.get(row.handle, ())

    def perform_row_action(self, row: Node, label: str) -> None:
        self.calls.append(f"row_action:{row.handle}:{label}")

    def alert_sheets(self) -> list[Node]:
        return list(self.sheets)

    def close_search_suggestions(self) -> int:
        closed, self.suggestions = self.suggestions, 0
        if closed:
            self.calls.append("close_suggestions")
        return closed

    def dismiss_known_popovers(self) -> int:
        dismissed, self.popovers = self.popovers, 0
        if dismissed:
            self.calls.append(f"dismiss_popovers:{dismissed}")
        return dismissed

    def visible_calls(self) -> list[str]:
        return [call for call in self.calls if call not in ("message_rows", "folder_rows")]

    def _open_folder(self, node: Node) -> None:
        parsed = parse_folder(node.description)
        if parsed is None:
            return
        self.folder = parsed.name
        if parsed.name in self.folder_rows_by_name:
            self.rows = list(self.folder_rows_by_name[parsed.name])

    def _marked(self, node: Node) -> Node:
        return replace(node, selected=node.handle in self.selected)
