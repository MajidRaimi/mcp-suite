from __future__ import annotations

from collections.abc import Callable

from office_mcp.features.outlook.compose.gateway import ComposeSurface, RecipientField
from office_mcp.shared.ax import AxApp
from office_mcp.shared.errors import ElementNotFoundError
from office_mcp.shared.tree import element_label, find, find_all

EDITOR = "Editor"
READING_PANE = "Reading Pane"
SUBJECT = "Subject"
SEND = "Send"
SEND_HELP = "Send the message"
DISCARD = "Discard"
DISCARD_HELP = "Discard this draft"
SENSITIVITY = "Sensitivity"
SENSITIVITY_HELP = "Set the sensitivity level of this message"
FIELD_LABELS = {
    RecipientField.TO: "To Recipients",
    RecipientField.CC: "Cc Recipients",
    RecipientField.BCC: "Bcc Recipients",
}
SHOW_BUTTONS = {RecipientField.CC: "Show Cc", RecipientField.BCC: "Show Bcc"}
TOKEN_ROLES = frozenset({"AXButton", "AXMenuButton", "AXStaticText"})
HEAVY_ROLES = frozenset({"AXTable", "AXOutline"})
SURFACE_DEPTH = 30
EDITOR_CLIMB = 8
BUTTON_DEPTH = 12
TOKEN_DEPTH = 5


class SurfaceFinder:
    def __init__(self, app: AxApp) -> None:
        self._app = app

    def inline_surface(self, main_window: object) -> ComposeSurface | None:
        editor = find(
            self._app, main_window, self._is_editor, SURFACE_DEPTH, prune=self._heavy_or_message
        )
        if editor is None:
            return None
        container: object | None = editor
        for _ in range(EDITOR_CLIMB):
            container = self._app.attribute(container, "AXParent") if container else None
            if container is None:
                return None
            if self.send_buttons(container):
                return ComposeSurface(handle=container, window=main_window, inline=True)
        return None

    def window_surface(self, window: object) -> ComposeSurface:
        return ComposeSurface(handle=window, window=window, inline=False)

    def send_buttons(self, root: object) -> list[object]:
        return self._buttons(root, self._is_send)

    def sensitivity_button(self, root: object) -> object | None:
        found = self._buttons(root, self._is_sensitivity)
        return found[0] if len(found) == 1 else None

    def shown_texts(self, root: object) -> list[str]:
        found = find_all(
            self._app,
            root,
            self._has_role("AXStaticText"),
            SURFACE_DEPTH,
            prune=self._heavy_or_editor,
        )
        texts = (
            self._app.text(element, "AXValue") or element_label(self._app, element)
            for element in found
        )
        return [text for text in texts if text.strip()]

    def discard_button(self, root: object) -> object:
        return self._single(self._buttons(root, self._is_discard), "Discard button")

    def body_area(self, surface: ComposeSurface) -> object:
        editor = find(self._app, surface.handle, self._is_editor, SURFACE_DEPTH, prune=self._heavy)
        if editor is None:
            raise ElementNotFoundError("the compose body editor was not found")
        area = find(self._app, editor, self._has_role("AXTextArea"), SURFACE_DEPTH)
        if area is None:
            raise ElementNotFoundError("the compose body text area was not found")
        return area

    def subject_field(self, surface: ComposeSurface) -> object:
        return self._text_field(surface, SUBJECT)

    def recipient_field(self, surface: ComposeSurface, field: RecipientField) -> object | None:
        return self._find_text_field(surface, FIELD_LABELS[field])

    def show_button(self, surface: ComposeSurface, field: RecipientField) -> object | None:
        label = SHOW_BUTTONS.get(field)
        if label is None:
            return None
        found = self._buttons(surface.handle, lambda e: element_label(self._app, e) == label)
        return found[0] if found else None

    def token_labels(self, field_element: object) -> list[tuple[str, tuple[str, ...]]]:
        container = self._app.attribute(field_element, "AXParent") or field_element
        tokens = find_all(
            self._app,
            container,
            lambda e: self._app.text(e, "AXRole") in TOKEN_ROLES,
            TOKEN_DEPTH,
            prune=lambda e: self._app.text(e, "AXRole") in TOKEN_ROLES,
        )
        return [
            (
                self._token_label(token),
                tuple(
                    self._app.text(token, name)
                    for name in ("AXTitle", "AXDescription", "AXValue", "AXHelp")
                ),
            )
            for token in tokens
            if not self._is_field_control(token)
        ]

    def element_alive(self, element: object) -> bool:
        return self._app.attribute(element, "AXRole") is not None

    def window_alive(self, window: object) -> bool:
        return any(self._app.same(window, candidate) for candidate in self._app.windows())

    def _token_label(self, element: object) -> str:
        return element_label(self._app, element) or " ".join(
            self._app.text(element, "AXValue").split()
        )

    def _is_field_control(self, element: object) -> bool:
        label = self._token_label(element)
        return label.startswith("Show ") or not label

    def _text_field(self, surface: ComposeSurface, label: str) -> object:
        found = self._find_text_field(surface, label)
        if found is None:
            raise ElementNotFoundError(f"the compose {label} field was not found")
        return found

    def _find_text_field(self, surface: ComposeSurface, label: str) -> object | None:
        return find(
            self._app,
            surface.handle,
            lambda e: (
                self._app.text(e, "AXRole") in ("AXTextField", "AXComboBox")
                and self._app.text(e, "AXDescription") == label
            ),
            SURFACE_DEPTH,
            prune=self._heavy,
        )

    def _buttons(self, root: object, matches: Callable[[object], bool]) -> list[object]:
        return find_all(
            self._app,
            root,
            lambda e: self._app.text(e, "AXRole") == "AXButton" and matches(e),
            BUTTON_DEPTH,
            prune=self._heavy,
        )

    def _single(self, found: list[object], label: str) -> object:
        if len(found) != 1:
            raise ElementNotFoundError(f"expected one {label} in the compose, found {len(found)}")
        return found[0]

    def _is_send(self, element: object) -> bool:
        return (
            self._app.text(element, "AXTitle") == SEND
            or self._app.text(element, "AXDescription") == SEND
            or self._app.text(element, "AXHelp") == SEND_HELP
        )

    def _is_discard(self, element: object) -> bool:
        return (
            self._app.text(element, "AXDescription") == DISCARD
            or self._app.text(element, "AXTitle") == DISCARD
            or self._app.text(element, "AXHelp") == DISCARD_HELP
        )

    def _is_sensitivity(self, element: object) -> bool:
        description = self._app.text(element, "AXDescription")
        return (
            self._app.text(element, "AXHelp") == SENSITIVITY_HELP
            or description == SENSITIVITY
            or description.startswith(f"{SENSITIVITY},")
        )

    def _is_editor(self, element: object) -> bool:
        return self._app.text(element, "AXRole") == "AXGroup" and (
            self._app.text(element, "AXDescription") == EDITOR
        )

    def _has_role(self, role: str) -> Callable[[object], bool]:
        return lambda element: self._app.text(element, "AXRole") == role

    def _heavy(self, element: object) -> bool:
        return self._app.text(element, "AXRole") in HEAVY_ROLES

    def _heavy_or_editor(self, element: object) -> bool:
        return self._heavy(element) or self._is_editor(element)

    def _heavy_or_message(self, element: object) -> bool:
        if self._heavy(element):
            return True
        return self._app.text(element, "AXRole") == "AXWebArea" and (
            self._app.text(element, "AXDescription") == READING_PANE
        )
