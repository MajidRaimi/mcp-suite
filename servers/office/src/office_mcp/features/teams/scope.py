from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from office_mcp.features.teams.header import open_thread_id
from office_mcp.features.teams.rail import TeamsApp, rail_button, selected_app
from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.errors import ElementNotFoundError, UiTimeoutError
from office_mcp.shared.focus import FocusOutcome, ForegroundSession
from office_mcp.shared.node import Node
from office_mcp.shared.scope import ScopeOptions

if TYPE_CHECKING:
    from office_mcp.features.teams.driver import TeamsDriver

RESTORE_ENHANCED = "restore the Teams accessibility mode"
GIVE_FOCUS_BACK = "give the front back"
RESTORE_VIEW = "restore the Teams view"
CLOSE_POPUPS = "close Teams menus"
CHAT_PANE_WAIT = 4.0


@dataclass(frozen=True)
class TeamsView:
    app: TeamsApp | None
    thread_id: str | None


class TeamsScope:
    def __init__(self, driver: TeamsDriver, options: ScopeOptions) -> None:
        self._driver = driver
        self._options = options
        self._session: ForegroundSession | None = None
        self.cleanup = CleanupStack()

    def __enter__(self) -> TeamsScope:
        self.cleanup.__enter__()
        try:
            self._enter()
        except BaseException as error:
            self.cleanup.__exit__(type(error), error, error.__traceback__)
            raise
        return self

    def __exit__(self, *exc: object) -> None:
        error = exc[1] if len(exc) > 1 and isinstance(exc[1], BaseException) else None
        self.cleanup.__exit__(type(error) if error is not None else None, error, None)

    def ensure_front(self) -> None:
        if self._session is None:
            self._start_session()
            return
        self._session.ensure_front()

    @property
    def focus(self) -> FocusOutcome:
        return self._session.outcome if self._session is not None else FocusOutcome.NOT_NEEDED

    def show_app(self, app: TeamsApp) -> Node:
        gateway = self._driver.gateway
        root = gateway.snapshot()
        if selected_app(root) is app:
            return root
        button = rail_button(root, app)
        if button is None:
            raise ElementNotFoundError(f"the Teams rail has no {app.name.title()} button")
        gateway.press(button)
        return self._driver.waiter.require(
            lambda: self._showing(app),
            lambda: UiTimeoutError(f"Teams did not switch to {app.name.title()} in time"),
        )

    def _enter(self) -> None:
        previous = self._driver.app.set_enhanced(True)
        self.cleanup.push(RESTORE_ENHANCED, lambda: self._restore_enhanced(previous))
        if self._options.foreground:
            self._start_session()
        self._driver.gateway.close_popups()
        if self._options.restore_view:
            view = self._capture()
            self.cleanup.push(RESTORE_VIEW, lambda: self._restore(view))
        self.cleanup.push(CLOSE_POPUPS, self._close_popups)

    def _start_session(self) -> None:
        driver = self._driver
        session = ForegroundSession(driver.workspace, driver.app, driver.waiter, driver.focus)
        session.__enter__()
        self._session = session
        self.cleanup.push(GIVE_FOCUS_BACK, session.__exit__)

    def _capture(self) -> TeamsView:
        root = self._driver.gateway.snapshot()
        app = selected_app(root)
        thread_id = open_thread_id(root) if app is TeamsApp.CHAT else None
        return TeamsView(app=app, thread_id=thread_id)

    def _restore(self, view: TeamsView) -> None:
        if view.app is not None:
            self.show_app(view.app)
        if view.thread_id is None:
            return
        if self._shown_thread() == view.thread_id:
            return
        row = self._driver.opener.find_self(view.thread_id)
        if row is None:
            raise ElementNotFoundError("the chat that was open before is no longer in the list")
        self._driver.opener.open_row(row)

    def _shown_thread(self) -> str | None:
        return self._driver.waiter.until(
            lambda: open_thread_id(self._driver.gateway.snapshot()), CHAT_PANE_WAIT
        )

    def _close_popups(self) -> None:
        self._driver.gateway.close_popups()

    def _showing(self, app: TeamsApp) -> Node | None:
        root = self._driver.gateway.snapshot()
        return root if selected_app(root) is app else None

    def _restore_enhanced(self, previous: bool | None) -> None:
        self._driver.app.set_enhanced(bool(previous))
