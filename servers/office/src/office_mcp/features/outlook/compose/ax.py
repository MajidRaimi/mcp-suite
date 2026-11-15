from __future__ import annotations

from collections.abc import Sequence
from contextlib import suppress

from office_mcp.features.outlook.compose.gateway import (
    AlertSheet,
    ComposeSurface,
    RecipientField,
    RecipientRead,
    ReplyKind,
)
from office_mcp.features.outlook.compose.recipients import parse_token, pending_text
from office_mcp.features.outlook.compose.sensitivity import require_label, shows_label
from office_mcp.features.outlook.compose.surface import SurfaceFinder
from office_mcp.features.outlook.gateway import OutlookGateway
from office_mcp.features.outlook.window import OutlookWindow, classify_window
from office_mcp.shared.ax import MENU_OPEN_CODES, AxApp, as_sequence
from office_mcp.shared.errors import (
    ComposeBusyError,
    ElementNotFoundError,
    OfficeError,
    UiTimeoutError,
)
from office_mcp.shared.events import Key, Modifier, Shortcut
from office_mcp.shared.menus import MenuBar, PopupMenus
from office_mcp.shared.node import Node
from office_mcp.shared.recipients import normalize_email
from office_mcp.shared.text_input import TextWriter
from office_mcp.shared.waiting import Waiter
from office_mcp.shared.windows import WindowLocator

NEW_MAIL = ("File", "New", "Mail")
NEW_MAIL_SHORTCUT = Shortcut(Key.N, Modifier.COMMAND)
UNDO_SEND = ("Edit", "Undo Send")
OPEN_TIMEOUT = 8.0
TOKEN_TIMEOUT = 5.0
FIELD_TIMEOUT = 3.0
ALERT_TIMEOUT = 3.0
CLOSE_TIMEOUT = 8.0
LABEL_TIMEOUT = 4.0
NO_SENSITIVITY = "the compose has no Sensitivity button, so no label can be set"
DISCARD = "Discard"
INLINE_BUSY = "a draft is already open in Outlook's reading pane; finish or discard it first"
TOKEN_SEPARATOR = ";"
COMMIT_TOKEN = Shortcut(Key.TAB)


class AxCompose:
    def __init__(
        self,
        app: AxApp,
        windows: WindowLocator,
        menus: MenuBar,
        popups: PopupMenus,
        text: TextWriter,
        outlook: OutlookGateway,
        waiter: Waiter,
    ) -> None:
        self._app = app
        self._windows = windows
        self._menus = menus
        self._popups = popups
        self._text = text
        self._outlook = outlook
        self._waiter = waiter
        self._finder = SurfaceFinder(app)

    def surfaces(self) -> list[ComposeSurface]:
        found = [
            self._finder.window_surface(probe.handle)
            for probe in self._windows.probes()
            if classify_window(probe) is OutlookWindow.COMPOSE
        ]
        inline = self._inline()
        return [*found, inline] if inline is not None else found

    def stray_surfaces(self, before: Sequence[ComposeSurface]) -> list[ComposeSurface]:
        return [surface for surface in self.surfaces() if not self._known(surface, before)]

    def open_new(self) -> ComposeSurface:
        if self._inline() is not None:
            raise ComposeBusyError(INLINE_BUSY)
        before = self.surfaces()
        self._menus.press(NEW_MAIL, NEW_MAIL_SHORTCUT)
        return self._waiter.require(
            lambda: next(iter(self.stray_surfaces(before)), None),
            lambda: UiTimeoutError("Outlook did not open a new mail in time"),
            OPEN_TIMEOUT,
        )

    def open_reply(self, row: Node, kind: ReplyKind) -> ComposeSurface:
        if self._inline() is not None:
            raise ComposeBusyError(INLINE_BUSY)
        before = self.surfaces()
        self._outlook.select(row)
        button = self._waiter.require(
            lambda: self._pane_button(kind),
            lambda: ElementNotFoundError(f"the reading pane shows no {kind.value!r} button"),
            OPEN_TIMEOUT,
        )
        self._outlook.press_button(button)
        return self._waiter.require(
            lambda: next(iter(self.stray_surfaces(before)), None),
            lambda: UiTimeoutError(f"Outlook did not open the {kind.value} compose in time"),
            OPEN_TIMEOUT,
        )

    def dismiss_popovers(self) -> int:
        return self._outlook.dismiss_known_popovers()

    def body(self, surface: ComposeSurface) -> str:
        return self._text.read(self._finder.body_area(surface))

    def insert_body(self, surface: ComposeSurface, text: str) -> None:
        self._text.write(self._finder.body_area(surface), text, at_start=True)

    def subject(self, surface: ComposeSurface) -> str:
        return self._app.text(self._finder.subject_field(surface), "AXValue")

    def set_subject(self, surface: ComposeSurface, text: str) -> None:
        self._text.set_value(self._finder.subject_field(surface), text)

    def add_recipient(self, surface: ComposeSurface, field: RecipientField, address: str) -> None:
        element = self._field(surface, field)
        self._text.write(element, f"{address}{TOKEN_SEPARATOR}")
        self._text.press_in(element, COMMIT_TOKEN)
        wanted = normalize_email(address)
        self._waiter.until(
            lambda: any(token.address == wanted for token in self._read(element).tokens),
            TOKEN_TIMEOUT,
        )

    def recipients(self, surface: ComposeSurface, field: RecipientField) -> RecipientRead:
        element = self._finder.recipient_field(surface, field)
        if element is None:
            return RecipientRead(tokens=(), pending="")
        return self._read(element)

    def set_sensitivity(self, surface: ComposeSurface, path: tuple[str, ...]) -> None:
        button = self._finder.sensitivity_button(surface.handle)
        if button is None:
            raise ElementNotFoundError(NO_SENSITIVITY)
        menu = self._popups.open(button, "AXPress")
        try:
            self._popups.choose(menu, path)
        except OfficeError:
            with suppress(OfficeError):
                self._popups.close(menu)
            raise
        self._waiter.until(
            lambda: any(shows_label(text, path[-1]) for text in self.shown_texts(surface)),
            LABEL_TIMEOUT,
        )
        require_label(self.shown_texts(surface), path)

    def shown_texts(self, surface: ComposeSurface) -> list[str]:
        return self._finder.shown_texts(surface.handle)

    def send_button(self, surface: ComposeSurface) -> Node:
        found = self._finder.send_buttons(surface.handle)
        if len(found) != 1:
            raise ElementNotFoundError(
                f"expected one Send button in the compose, found {len(found)}"
            )
        return Node(role="AXButton", title="Send", handle=found[0])

    def press_send(self, button: Node) -> None:
        if button.handle is None:
            raise ElementNotFoundError("the Send button has no element")
        self._app.perform(button.handle, "AXPress", "Send", tolerate=MENU_OPEN_CODES)

    def is_open(self, surface: ComposeSurface) -> bool:
        if not self._finder.window_alive(surface.window):
            return False
        if not self._finder.element_alive(surface.handle):
            return False
        return bool(self._finder.send_buttons(surface.handle))

    def alerts(self, window: object) -> list[AlertSheet]:
        return [
            to_alert(self._app.snapshot(child))
            for child in as_sequence(self._app.attribute(window, "AXChildren"))
            if self._app.text(child, "AXRole") == "AXSheet"
        ]

    def new_alert(self, window: object, before: Sequence[AlertSheet]) -> AlertSheet | None:
        return next(
            (
                alert
                for alert in self.alerts(window)
                if not any(self._same_sheet(alert, old) for old in before)
            ),
            None,
        )

    def press_alert(self, alert: AlertSheet, title: str) -> None:
        button = alert.node.find(lambda node: node.role == "AXButton" and node.label() == title)
        if button is None:
            raise ElementNotFoundError(f"the alert has no {title!r} button")
        self._outlook.press_button(button)

    def discard(self, surface: ComposeSurface) -> None:
        if not self.is_open(surface):
            return
        self._raise(surface)
        button = self._finder.discard_button(surface.handle)
        self._app.perform(button, "AXPress", DISCARD, tolerate=MENU_OPEN_CODES)
        alert = self._waiter.until(lambda: self._discard_alert(surface.window), ALERT_TIMEOUT)
        if alert is not None:
            self.press_alert(alert, DISCARD)
        self._waiter.require(
            lambda: not self.is_open(surface),
            lambda: UiTimeoutError("the unsent draft did not close after Discard"),
            CLOSE_TIMEOUT,
        )

    def undo_send_enabled(self) -> bool:
        return self._menus.entry(UNDO_SEND).enabled

    def undo_send(self) -> None:
        self._menus.press(UNDO_SEND)

    def _discard_alert(self, window: object) -> AlertSheet | None:
        return next((alert for alert in self.alerts(window) if DISCARD in alert.buttons), None)

    def _same_sheet(self, first: AlertSheet, second: AlertSheet) -> bool:
        if first.node.handle is None or second.node.handle is None:
            return False
        return self._app.same(first.node.handle, second.node.handle)

    def _raise(self, surface: ComposeSurface) -> None:
        with suppress(OfficeError):
            self._app.set_attribute(surface.window, "AXMain", True)
        self._app.perform(surface.window, "AXRaise", "compose window", tolerate=MENU_OPEN_CODES)

    def _field(self, surface: ComposeSurface, field: RecipientField) -> object:
        element = self._finder.recipient_field(surface, field)
        if element is not None:
            return element
        show = self._finder.show_button(surface, field)
        if show is None:
            raise ElementNotFoundError(f"the compose has no {field.value} field to show")
        self._app.perform(show, "AXPress", f"Show {field.value}", tolerate=MENU_OPEN_CODES)
        return self._waiter.require(
            lambda: self._finder.recipient_field(surface, field),
            lambda: ElementNotFoundError(f"the compose {field.value} field did not appear"),
            FIELD_TIMEOUT,
        )

    def _read(self, element: object) -> RecipientRead:
        tokens = tuple(
            parse_token(label, extra) for label, extra in self._finder.token_labels(element)
        )
        return RecipientRead(tokens=tokens, pending=pending_text(self._text.read(element)))

    def _inline(self) -> ComposeSurface | None:
        return self._finder.inline_surface(self._outlook.main_probe().handle)

    def _pane_button(self, kind: ReplyKind) -> Node | None:
        pane = self._outlook.reading_pane()
        if pane is None:
            return None
        return pane.find(lambda node: node.role == "AXButton" and node.title == kind.value)

    def _known(self, surface: ComposeSurface, before: Sequence[ComposeSurface]) -> bool:
        return any(
            old.inline == surface.inline and self._app.same(old.handle, surface.handle)
            for old in before
        )


def to_alert(sheet: Node) -> AlertSheet:
    buttons = tuple(
        node.label() for node in sheet.walk() if node.role == "AXButton" and node.label()
    )
    return AlertSheet(text=sheet.text() or sheet.label(), buttons=buttons, node=sheet)
