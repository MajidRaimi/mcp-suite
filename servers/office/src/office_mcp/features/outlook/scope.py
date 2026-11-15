from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from office_mcp.features.outlook.view import ViewState, capture_view, restore_view
from office_mcp.shared.cleanup import CleanupStack
from office_mcp.shared.focus import FocusOutcome, ForegroundSession
from office_mcp.shared.scope import ScopeOptions

if TYPE_CHECKING:
    from office_mcp.features.outlook.driver import OutlookDriver

RESTORE_VIEW = "restore the Outlook view"
GIVE_FOCUS_BACK = "give the front back"


class OutlookScope:
    def __init__(self, driver: OutlookDriver, options: ScopeOptions) -> None:
        self._driver = driver
        self._options = options
        self._session: ForegroundSession | None = None
        self._state: ViewState | None = None
        self._unbind: Callable[[], None] = lambda: None
        self.cleanup = CleanupStack()
        self.view_notes: tuple[str, ...] = ()

    def __enter__(self) -> OutlookScope:
        self.cleanup.__enter__()
        try:
            self._driver.health()
            if self._options.restore_view:
                self._state = capture_view(self._driver.gateway)
                self.cleanup.push(RESTORE_VIEW, self._restore_once)
            if self._options.foreground:
                self._start_session()
                self._driver.popups.close()
        except BaseException as error:
            self.cleanup.__exit__(type(error), error, error.__traceback__)
            raise
        self._unbind = self._driver.front.bind(self.ensure_front)
        return self

    def __exit__(self, *exc: object) -> None:
        error = exc[1] if len(exc) > 1 and isinstance(exc[1], BaseException) else None
        try:
            self.cleanup.__exit__(type(error) if error is not None else None, error, None)
        finally:
            self._unbind()

    def ensure_front(self) -> None:
        if self._session is None:
            self._start_session()
            return
        self._session.ensure_front()

    @property
    def focus(self) -> FocusOutcome:
        return self._session.outcome if self._session is not None else FocusOutcome.NOT_NEEDED

    def _start_session(self) -> None:
        driver = self._driver
        session = ForegroundSession(driver.workspace, driver.app, driver.waiter, driver.focus)
        session.__enter__()
        self._session = session
        self.cleanup.push(GIVE_FOCUS_BACK, lambda: self._give_back(session))

    def _give_back(self, session: ForegroundSession) -> None:
        try:
            self._restore_once()
        finally:
            session.__exit__(None, None, None)

    def _restore_once(self) -> None:
        state, self._state = self._state, None
        if state is not None:
            self._restore(state)

    def _restore(self, state: ViewState) -> None:
        self.view_notes = restore_view(self._driver.gateway, state, self._driver.waiter)
