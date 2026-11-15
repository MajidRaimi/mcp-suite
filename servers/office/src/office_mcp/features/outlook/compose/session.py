from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from types import TracebackType

from office_mcp.features.outlook.compose.gateway import (
    AlertSheet,
    ComposeGateway,
    ComposeSurface,
    RecipientField,
    RecipientRead,
)
from office_mcp.features.outlook.compose.sensitivity import require_label
from office_mcp.features.outlook.compose.verify import (
    body_is_free,
    require_free_body,
    safe_alert_button,
    written_part,
)
from office_mcp.shared.errors import (
    CleanupError,
    ComposeBusyError,
    ComposeError,
    OfficeError,
    SendRefusedError,
    error_text,
)
from office_mcp.shared.focus import FocusOutcome
from office_mcp.shared.scope import ScopeFactory, ScopeOptions, UiScope
from office_mcp.shared.waiting import Waiter

AFTER_OPEN = "after_open"
AFTER_RECIPIENTS = "after_recipients"
AFTER_SUBJECT = "after_subject"
AFTER_BODY = "after_body"
BEFORE_SEND = "before_send"
FAIL_POINTS = (AFTER_OPEN, AFTER_RECIPIENTS, AFTER_SUBJECT, AFTER_BODY, BEFORE_SEND)
CLOSE_TIMEOUT = 10.0
DISCARDED_NOTE = "nothing was sent; the unsent draft was discarded and a copy is in Deleted Items"
NOT_DISCARDED_NOTE = (
    "nothing was sent, but the unsent draft could not be discarded ({reason}); close it in "
    "Outlook without sending"
)
KEPT_NOTE = (
    "nothing was sent and nothing was typed; the compose was left open so its existing text is kept"
)
ALERT_LEFT_OPEN = "Outlook's alert is still open; close it in Outlook without sending"
FOREIGN_ALERT_NOTE = (
    "an Outlook alert that this send did not open is showing; it was left untouched"
)
STRAY_LEFT_NOTE = (
    "a compose that opened late was left open because it could not be proven to be this "
    "server's; check it in Outlook and close it without sending if it is not yours"
)


@dataclass(frozen=True)
class Scoped[T]:
    value: T
    focus: FocusOutcome
    notes: tuple[str, ...]


def run_scoped[T](
    scope: ScopeFactory, options: ScopeOptions, work: Callable[[UiScope], T]
) -> Scoped[T]:
    scopes: list[UiScope] = []
    values: list[T] = []
    try:
        with scope(options) as ui:
            scopes.append(ui)
            values.append(work(ui))
    except CleanupError as error:
        if not values:
            raise
        return Scoped(values[0], scopes[0].focus, (error_text(error),))
    return Scoped(values[0], scopes[0].focus, ())


@dataclass(frozen=True)
class Observation:
    closed: bool
    notes: tuple[str, ...] = ()


class ComposeSession:
    def __init__(
        self,
        gateway: ComposeGateway,
        open_surface: Callable[[], ComposeSurface],
        waiter: Waiter,
        fail_at: str | None = None,
        keep_busy: bool = False,
    ) -> None:
        if fail_at is not None and fail_at not in FAIL_POINTS:
            raise ComposeError(f"fail_at must be one of {', '.join(FAIL_POINTS)}")
        self._gateway = gateway
        self._open_surface = open_surface
        self._waiter = waiter
        self._fail_at = fail_at
        self._keep_busy = keep_busy
        self._surface: ComposeSurface | None = None
        self._baseline = ""
        self._committed = False
        self._press_note = ""
        self._alerts_before: tuple[AlertSheet, ...] = ()

    def __enter__(self) -> ComposeSession:
        before = self._gateway.surfaces()
        try:
            self._surface = self._open_surface()
        except OfficeError as error:
            self._discard_strays(before, error)
            raise
        try:
            self._gateway.dismiss_popovers()
            self.checkpoint(AFTER_OPEN)
        except BaseException as error:
            self._discard_after(error)
            raise
        return self

    def __exit__(
        self,
        kind: type[BaseException] | None,
        error: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        if error is None or self._committed:
            return
        if self._keep_busy and isinstance(error, ComposeBusyError):
            error.add_note(KEPT_NOTE)
            return
        self._discard_after(error)

    @property
    def surface(self) -> ComposeSurface:
        if self._surface is None:
            raise ComposeError("the compose window is not open")
        return self._surface

    @property
    def committed(self) -> bool:
        return self._committed

    def checkpoint(self, name: str) -> None:
        if name == self._fail_at:
            raise ComposeError(f"stopped on purpose at {name}")

    def require_empty_body(self) -> None:
        baseline = self._gateway.body(self.surface)
        require_free_body(baseline)
        self._baseline = baseline

    def fill(
        self, recipients: Sequence[tuple[RecipientField, tuple[str, ...]]], subject: str, body: str
    ) -> None:
        for field, addresses in recipients:
            self.add_recipients(field, addresses)
        self.checkpoint(AFTER_RECIPIENTS)
        if subject:
            self.set_subject(subject)
        self.checkpoint(AFTER_SUBJECT)
        self.insert_body(body)
        self.checkpoint(AFTER_BODY)

    def add_recipients(self, field: RecipientField, addresses: tuple[str, ...]) -> None:
        for address in addresses:
            self._gateway.add_recipient(self.surface, field, address)

    def recipients(self, field: RecipientField) -> RecipientRead:
        return self._gateway.recipients(self.surface, field)

    def set_subject(self, text: str) -> None:
        self._gateway.set_subject(self.surface, text)

    def subject(self) -> str:
        return self._gateway.subject(self.surface)

    def insert_body(self, text: str) -> None:
        if text:
            self._gateway.insert_body(self.surface, text)

    def set_sensitivity(self, path: tuple[str, ...]) -> None:
        if path:
            self._gateway.set_sensitivity(self.surface, path)

    def require_sensitivity(self, path: tuple[str, ...]) -> None:
        if path:
            require_label(self._gateway.shown_texts(self.surface), path)

    def written_body(self) -> str:
        return written_part(self._gateway.body(self.surface), self._baseline)

    def commit(self) -> None:
        if self._committed:
            raise ComposeError("Send was already pressed for this message")
        self.checkpoint(BEFORE_SEND)
        button = self._gateway.send_button(self.surface)
        self._alerts_before = tuple(self._gateway.alerts(self.surface.window))
        self._committed = True
        try:
            self._gateway.press_send(button)
        except OfficeError as error:
            self._press_note = f"pressing Send reported: {error_text(error)}"

    def observe(self, timeout: float = CLOSE_TIMEOUT) -> Observation:
        if not self._committed:
            raise ComposeError("Send was not pressed")
        notes = (self._press_note,) if self._press_note else ()
        self._waiter.until(self._settled, timeout)
        refusal = self._refusal()
        if refusal is not None:
            self._refuse(refusal)
        if self._gateway.is_open(self.surface):
            return Observation(closed=False, notes=notes)
        if self._new_alert() is not None:
            notes = (*notes, FOREIGN_ALERT_NOTE)
        return Observation(closed=True, notes=notes)

    def discard(self) -> None:
        self._gateway.discard(self.surface)

    def _settled(self) -> bool:
        return self._refusal() is not None or not self._gateway.is_open(self.surface)

    def _new_alert(self) -> AlertSheet | None:
        return self._gateway.new_alert(self.surface.window, self._alerts_before)

    def _refusal(self) -> AlertSheet | None:
        alert = self._new_alert()
        if alert is None or not self._gateway.is_open(self.surface):
            return None
        return alert

    def _refuse(self, alert: AlertSheet) -> None:
        refusal = SendRefusedError(
            f"Outlook stopped the send with an alert, so nothing was sent: {alert.text[:300]!r}"
        )
        choice = safe_alert_button(alert.buttons)
        if choice is None:
            refusal.add_note(ALERT_LEFT_OPEN)
            raise refusal
        try:
            self._gateway.press_alert(alert, choice)
            self.discard()
        except OfficeError as failure:
            refusal.add_note(NOT_DISCARDED_NOTE.format(reason=error_text(failure)))
            raise refusal from failure
        refusal.add_note(DISCARDED_NOTE)
        raise refusal

    def _discard_after(self, error: BaseException) -> None:
        if self._surface is None:
            return
        try:
            self.discard()
        except OfficeError as failure:
            error.add_note(NOT_DISCARDED_NOTE.format(reason=error_text(failure)))
            return
        error.add_note(DISCARDED_NOTE)

    def _discard_strays(self, before: list[ComposeSurface], error: OfficeError) -> None:
        strays = self._gateway.stray_surfaces(before)
        if len(strays) == 1 and self._is_blank(strays[0]):
            self._surface = strays[0]
            self._discard_after(error)
            self._surface = None
            return
        if strays:
            error.add_note(STRAY_LEFT_NOTE)

    def _is_blank(self, surface: ComposeSurface) -> bool:
        try:
            if not body_is_free(self._gateway.body(surface)):
                return False
            if self._gateway.subject(surface).strip():
                return False
            return all(
                is_empty(self._gateway.recipients(surface, field)) for field in RecipientField
            )
        except OfficeError:
            return False


def is_empty(read: RecipientRead) -> bool:
    return not read.tokens and not read.pending.strip()
