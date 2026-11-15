from __future__ import annotations

import pytest

from office_mcp.features.outlook.window import (
    OutlookMode,
    OutlookWindow,
    account_from_title,
    classify_window,
    detect_mode,
    folder_from_title,
)
from office_mcp.shared.windows import WindowProbe

SWITCHER = frozenset({"Mail", "Calendar", "People", "To Do"})
MAIL_LANDMARKS = SWITCHER | {"Message List", "Toolbar", "New Mail; Press space for other actions"}
COMPOSE_LANDMARKS = frozenset({"Send", "Send the message", "Discard", "To Recipients", "Subject"})


def probe(title: str, landmarks: frozenset[str], subrole: str = "AXStandardWindow") -> WindowProbe:
    return WindowProbe(
        handle=title, title=title, subrole=subrole, main=True, focused=True, landmarks=landmarks
    )


@pytest.mark.parametrize(
    ("title", "landmarks", "kind"),
    [
        ("Inbox • a@example.com", MAIL_LANDMARKS, OutlookWindow.MAIN),
        ('Searching "a@example.com"', MAIL_LANDMARKS | {"Cancel Search"}, OutlookWindow.MAIN),
        (
            "Calendar",
            SWITCHER | {"Calendar View, currently selected item Week"},
            OutlookWindow.MAIN,
        ),
        ("Untitled • a@example.com", COMPOSE_LANDMARKS, OutlookWindow.COMPOSE),
        ("Inbox • a@example.com", MAIL_LANDMARKS | {"Send", "Discard"}, OutlookWindow.MAIN),
        ("menuWindow", frozenset({"Current Folder"}), OutlookWindow.OTHER),
    ],
)
def test_classifies_outlook_windows(
    title: str, landmarks: frozenset[str], kind: OutlookWindow
) -> None:
    assert classify_window(probe(title, landmarks)) is kind


def test_send_button_outside_a_standard_window_is_not_a_compose_window() -> None:
    assert classify_window(probe("", COMPOSE_LANDMARKS, subrole="AXUnknown")) is OutlookWindow.OTHER


@pytest.mark.parametrize(
    ("title", "landmarks", "mode"),
    [
        ("Inbox • a@example.com", MAIL_LANDMARKS, OutlookMode.MAIL),
        ('Searching "a@example.com"', MAIL_LANDMARKS, OutlookMode.SEARCH),
        ("Inbox • a@example.com", MAIL_LANDMARKS | {"Cancel Search"}, OutlookMode.SEARCH),
        ("Calendar", SWITCHER, OutlookMode.CALENDAR),
        ("Week", SWITCHER | {"Calendar View, currently selected item Week"}, OutlookMode.CALENDAR),
        ("Sent Items • a@example.com", SWITCHER, OutlookMode.MAIL),
        ("People", SWITCHER, OutlookMode.OTHER),
    ],
)
def test_detects_the_mode_from_title_and_landmarks(
    title: str, landmarks: frozenset[str], mode: OutlookMode
) -> None:
    assert detect_mode(probe(title, landmarks)) is mode


@pytest.mark.parametrize(
    ("title", "account"),
    [
        ("Inbox • a@example.com", "a@example.com"),
        ('Searching "a@example.com"', "a@example.com"),
        ("Untitled • a@example.com", "a@example.com"),
        ("Calendar", None),
        ("Inbox • not an address", None),
    ],
)
def test_reads_the_account_from_the_title(title: str, account: str | None) -> None:
    assert account_from_title(title) == account


def test_reads_the_folder_from_the_title() -> None:
    assert folder_from_title("Sent Items • a@example.com") == "Sent Items"
    assert folder_from_title('Searching "a@example.com"') == ""
