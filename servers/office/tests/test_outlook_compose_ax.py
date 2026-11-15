from __future__ import annotations

from dataclasses import dataclass

import pytest
from fakes.ax_backend import FakeElement, element
from fakes.events import RecordingEvents
from fakes.pasteboard import FakePasteboard
from fakes.trees import AppRig, app_rig, menu_bar, menu_item, outlook_main_window, window

from office_mcp.features.outlook.ax import AxOutlook
from office_mcp.features.outlook.compose.ax import AxCompose
from office_mcp.features.outlook.compose.gateway import ComposeSurface, RecipientField
from office_mcp.shared.errors import ComposeBusyError, ElementNotFoundError, TextEntryError
from office_mcp.shared.events import Key, Shortcut
from office_mcp.shared.menus import MenuBar, PopupMenus
from office_mcp.shared.text_input import TextWriter
from office_mcp.shared.windows import WindowLocator

SIGNATURE = "\nJane Doe"
LABEL_SHOWN: dict[str, str] = {}
AX_FAILURE = -25200


@dataclass
class ComposeApp:
    rig: AppRig
    compose: AxCompose
    windows: list[FakeElement]
    events: RecordingEvents


def recipient_box(label: str) -> tuple[FakeElement, FakeElement]:
    field = element("AXTextField", AXDescription=label, AXValue="")
    box = element("AXScrollArea", field)

    def tokenize(value: object) -> None:
        text = str(value).rstrip(";")
        box.add(element("AXButton", AXDescription=f"Bob Sample, {text}"))

    field.on_set["AXSelectedText"] = tokenize
    return field, box


def compose_window(app: ComposeApp) -> FakeElement:
    area = element("AXTextArea", AXValue=SIGNATURE)

    def insert(value: object) -> None:
        area.attributes["AXValue"] = f"{value}{area.attributes['AXValue']}"

    area.on_set["AXSelectedText"] = insert
    _, to_box = recipient_box("To Recipients")
    show_cc = element("AXButton", AXDescription="Show Cc")
    discard = element("AXButton", AXDescription="Discard", AXHelp="Discard this draft")
    send = element("AXButton", AXDescription="Send", AXHelp="Send the message")
    sensitivity = element(
        "AXButton", AXDescription="Sensitivity", AXHelp="Set the sensitivity level of this message"
    )
    toolbar = element("AXToolbar", send, discard, sensitivity)
    editor = element("AXGroup", element("AXScrollArea", element("AXWebArea", area)))
    editor.attributes["AXDescription"] = "Editor"
    compose = window(
        "Untitled • jane@example.com",
        toolbar,
        to_box,
        show_cc,
        element("AXTextField", AXDescription="Subject", AXValue=""),
        editor,
        AXSubrole="AXStandardWindow",
    )

    def reveal_cc() -> int:
        compose.add(recipient_box("Cc Recipients")[1])
        return 0

    show_cc.actions["AXPress"] = reveal_cc

    def ask_to_discard() -> int:
        confirm = element("AXButton", AXTitle="Discard")
        confirm.actions["AXPress"] = lambda: close(app, compose)
        compose.add(
            element(
                "AXSheet",
                element("AXStaticText", AXValue="Discard draft message?"),
                element("AXButton", AXTitle="Continue Editing"),
                confirm,
            )
        )
        return 0

    discard.actions["AXPress"] = ask_to_discard
    send.action_codes["AXPress"] = AX_FAILURE
    sensitivity.actions["AXPress"] = lambda: open_labels(compose, sensitivity)
    return compose


def open_labels(compose: FakeElement, button: FakeElement) -> int:
    general = menu_item("General")
    internal = menu_item("Internal")
    menu = element("AXMenu", general, menu_item("Confidential", internal))

    def pick(label: str) -> int:
        menu.remove()
        if label:
            compose.add(element("AXStaticText", AXValue=label, AXHelp="Label policy"))
        return 0

    menu.actions["AXCancel"] = lambda: pick("")
    general.actions["AXPress"] = lambda: pick(LABEL_SHOWN.get("General", "General"))
    internal.actions["AXPress"] = lambda: pick(LABEL_SHOWN.get("Internal", "Internal"))
    button.add(menu)
    return 0


def close(app: ComposeApp, compose: FakeElement) -> int:
    app.windows.remove(compose)
    compose.remove()
    app.rig.root.attributes["AXWindows"] = list(app.windows)
    return 0


def build() -> ComposeApp:
    main = outlook_main_window()
    new_mail = menu_item("Mail", char="N", modifiers=0)
    other_mail = menu_item("Mail", char="N", modifiers=2)
    bar = menu_bar(
        ("File", [menu_item("New", new_mail, other_mail)]),
        ("Edit", [menu_item("Undo Send", enabled=False, virtual_key=36, modifiers=1)]),
    )
    rig = app_rig(main, menu_bar=bar)
    events = RecordingEvents()
    waiter = rig.app.waiter
    windows = WindowLocator(rig.app)
    outlook = AxOutlook(rig.app, windows, PopupMenus(rig.app, events, waiter), events)
    text = TextWriter(rig.app, events, FakePasteboard(), waiter)
    popups = PopupMenus(rig.app, events, waiter)
    compose = AxCompose(rig.app, windows, MenuBar(rig.app), popups, text, outlook, waiter)
    app = ComposeApp(rig, compose, [main], events)

    def open_compose() -> int:
        created = compose_window(app)
        rig.root.add(created)
        app.windows.append(created)
        rig.root.attributes["AXWindows"] = list(app.windows)
        return 0

    def wrong_item() -> int:
        pytest.fail("pressed the wrong Mail item")

    new_mail.actions["AXPress"] = open_compose
    other_mail.actions["AXPress"] = wrong_item
    return app


def opened() -> tuple[ComposeApp, ComposeSurface]:
    app = build()
    return app, app.compose.open_new()


def test_open_new_uses_the_command_n_item_and_finds_the_window() -> None:
    app, surface = opened()

    assert surface.inline is False
    assert surface.handle is app.windows[-1]
    assert app.compose.is_open(surface)
    assert len(app.compose.surfaces()) == 1


def test_recipient_is_typed_with_a_separator_and_read_back_as_a_token() -> None:
    app, surface = opened()

    app.compose.add_recipient(surface, RecipientField.TO, "bob@example.com")
    read = app.compose.recipients(surface, RecipientField.TO)

    assert [token.address for token in read.tokens] == ["bob@example.com"]
    assert read.pending == ""
    assert "bob@example.com;" in app.rig.backend.sets_of("AXSelectedText")
    assert app.compose.recipients(surface, RecipientField.BCC).tokens == ()


def test_cc_field_is_revealed_before_typing() -> None:
    app, surface = opened()

    app.compose.add_recipient(surface, RecipientField.CC, "alex@example.com")

    read = app.compose.recipients(surface, RecipientField.CC)
    assert [token.address for token in read.tokens] == ["alex@example.com"]


def test_body_goes_above_the_signature() -> None:
    app, surface = opened()

    app.compose.insert_body(surface, "Hello")

    assert app.compose.body(surface) == "Hello" + SIGNATURE
    assert (0, 0) in app.rig.backend.sets_of("AXSelectedTextRange")


def test_send_press_tolerates_the_misleading_error_code() -> None:
    app, surface = opened()

    button = app.compose.send_button(surface)
    app.compose.press_send(button)

    assert len(app.rig.backend.actions_of("AXPress")) >= 2


def test_discard_confirms_the_alert_and_waits_for_the_window_to_close() -> None:
    app, surface = opened()

    app.compose.discard(surface)

    assert not app.compose.is_open(surface)
    assert app.compose.surfaces() == []


def test_undo_send_reports_a_disabled_item() -> None:
    app = build()

    assert app.compose.undo_send_enabled() is False


def test_two_send_buttons_are_refused() -> None:
    app, surface = opened()
    app.windows[-1].add(element("AXButton", AXTitle="Send"))

    with pytest.raises(ElementNotFoundError, match="found 2"):
        app.compose.send_button(surface)


def inline_build() -> tuple[ComposeApp, FakeElement]:
    app = build()
    main = app.windows[0]
    pane = element("AXGroup", AXDescription="Reading Pane")
    main.add(pane)

    def open_inline() -> int:
        parts = list(compose_window(app).children)
        pane.add(element("AXGroup", *parts))
        return 0

    menu = app.rig.root.attributes["AXMenuBar"]
    assert isinstance(menu, FakeElement)
    new_mail = next(
        node
        for node in walk(menu)
        if node.attributes.get("AXTitle") == "Mail"
        and node.attributes.get("AXMenuItemCmdModifiers") == 0
    )
    new_mail.actions["AXPress"] = open_inline
    return app, main


def walk(node: FakeElement) -> list[FakeElement]:
    return [node, *(found for child in node.children for found in walk(child))]


def test_open_new_finds_a_new_mail_opened_inline_in_the_main_window() -> None:
    app, main = inline_build()

    surface = app.compose.open_new()

    assert surface.inline is True
    assert surface.window is main
    assert app.compose.is_open(surface)


def test_open_new_refuses_while_an_inline_draft_is_open() -> None:
    app, _ = inline_build()
    app.compose.open_new()

    with pytest.raises(ComposeBusyError, match="reading pane"):
        app.compose.open_new()


def test_a_recipient_typed_as_text_becomes_a_static_token_after_tab() -> None:
    app, surface = opened()
    compose = app.windows[-1]
    field = next(
        node for node in walk(compose) if node.attributes.get("AXDescription") == "To Recipients"
    )
    box = field.parent
    assert box is not None

    def type_text(value: object) -> None:
        field.attributes["AXValue"] = f"{field.attributes.get('AXValue', '')}{value}"

    def commit(_pid: int, shortcut: Shortcut) -> None:
        if shortcut.key is Key.TAB:
            field.attributes["AXValue"] = "\ufffc"
            box.add(element("AXStaticText", AXValue="bob@example.com, Available"))

    field.on_set["AXSelectedText"] = type_text
    app.events.on_press = commit

    app.compose.add_recipient(surface, RecipientField.TO, "bob@example.com")

    read = app.compose.recipients(surface, RecipientField.TO)
    assert [token.address for token in read.tokens] == ["bob@example.com"]
    assert read.pending == ""
    assert Shortcut(Key.TAB) in app.events.shortcuts()


def test_sensitivity_label_is_picked_from_the_menu_and_read_back() -> None:
    app, surface = opened()

    app.compose.set_sensitivity(surface, ("General",))

    assert "General" in app.compose.shown_texts(surface)


def test_nested_sensitivity_label_opens_its_submenu() -> None:
    app, surface = opened()

    app.compose.set_sensitivity(surface, ("Confidential", "Internal"))

    assert "Internal" in app.compose.shown_texts(surface)


def test_unknown_sensitivity_label_closes_the_menu_and_fails() -> None:
    app, surface = opened()

    with pytest.raises(ElementNotFoundError, match="Secret"):
        app.compose.set_sensitivity(surface, ("Secret",))

    roles = [node.attributes.get("AXRole") for node in walk(app.windows[-1])]
    assert "AXMenu" not in roles


def test_sensitivity_label_that_does_not_show_is_refused(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setitem(LABEL_SHOWN, "General", "")
    app, surface = opened()

    with pytest.raises(TextEntryError, match="General"):
        app.compose.set_sensitivity(surface, ("General",))


def test_compose_without_a_sensitivity_button_says_so() -> None:
    app, surface = opened()
    button = next(
        node
        for node in walk(app.windows[-1])
        if node.attributes.get("AXDescription") == "Sensitivity"
    )
    button.remove()

    with pytest.raises(ElementNotFoundError, match="no Sensitivity button"):
        app.compose.set_sensitivity(surface, ("General",))
