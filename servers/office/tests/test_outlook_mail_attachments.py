from __future__ import annotations

from pathlib import Path

import pytest
from fakes.ax_backend import FakeElement, element
from fakes.clock import FakeClock
from fakes.p1_mail import MailGateway, MailRig, PaneSpec, attachment_group, mail_rig, mail_row
from fakes.p1_mail import reading_pane as pane_for
from fakes.trees import menu_item, outlook_rows

from office_mcp.features.outlook.mail.attachments import AttachmentService
from office_mcp.features.outlook.mail.downloads import (
    PathDownloadsWatcher,
    download_pattern,
    move_into,
    safe_name,
    size_matches,
    unique_path,
)
from office_mcp.shared.errors import AttachmentError, TargetError, UiTimeoutError

CONTENT = b"a" * 12800


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("budget.xlsx", "budget.xlsx"),
        ("../../etc/passwd", "passwd"),
        ("..\\..\\notes.txt", "notes.txt"),
        (".hidden.pdf", "hidden.pdf"),
        ("re\x00port\x1f.pdf", "report.pdf"),
        ("a:b.txt", "ab.txt"),
        ("...", "attachment"),
        ("", "attachment"),
    ],
)
def test_safe_name(raw: str, expected: str) -> None:
    assert safe_name(raw) == expected


def test_unique_path_adds_a_numbered_suffix(tmp_path: Path) -> None:
    (tmp_path / "budget.xlsx").write_bytes(b"1")
    (tmp_path / "budget (1).xlsx").write_bytes(b"2")

    assert unique_path(tmp_path, "budget.xlsx") == tmp_path / "budget (2).xlsx"
    assert unique_path(tmp_path, "notes.txt") == tmp_path / "notes.txt"


def test_move_into_keeps_files_inside_the_directory(tmp_path: Path) -> None:
    source = tmp_path / "incoming.bin"
    source.write_bytes(b"data")
    target_dir = tmp_path / "saved"
    (target_dir).mkdir()
    (target_dir / "passwd").write_bytes(b"old")

    saved = move_into(source, target_dir, "../../passwd")

    assert saved == target_dir / "passwd (1)"
    assert saved.read_bytes() == b"data"
    assert not source.exists()


@pytest.mark.parametrize(
    ("shown", "actual", "ok"),
    [
        ("12.5 KB", 12800, True),
        ("12.5 KB", 12500, True),
        ("12.5 KB", 40000, False),
        ("2 MB", 2 * 1024 * 1024, True),
        ("300 bytes", 300, True),
        ("", 10, True),
        ("", 0, False),
    ],
)
def test_size_matches(shown: str, actual: int, ok: bool) -> None:
    assert size_matches(shown, actual) is ok


def test_download_pattern_accepts_numbered_copies() -> None:
    pattern = download_pattern("budget.xlsx")

    assert pattern.match("budget.xlsx")
    assert pattern.match("budget (3).xlsx")
    assert not pattern.match("budget.xlsx.download")
    assert not pattern.match("old budget.xlsx")


def test_watcher_waits_for_a_stable_completed_file(tmp_path: Path) -> None:
    clock = FakeClock()
    watcher = PathDownloadsWatcher(tmp_path, clock, clock.sleep)
    (tmp_path / "budget.xlsx").write_bytes(b"old copy")
    before = watcher.snapshot()
    (tmp_path / "budget (1).xlsx").write_bytes(CONTENT)

    found = watcher.wait_for("budget.xlsx", before, 30)

    assert found == tmp_path / "budget (1).xlsx"
    assert clock.now >= 1.0


def test_watcher_ignores_files_still_downloading(tmp_path: Path) -> None:
    clock = FakeClock()
    watcher = PathDownloadsWatcher(tmp_path, clock, clock.sleep)
    before = watcher.snapshot()
    (tmp_path / "budget.xlsx").write_bytes(CONTENT)
    (tmp_path / "budget.xlsx.download").mkdir()

    with pytest.raises(UiTimeoutError, match="did not finish downloading"):
        watcher.wait_for("budget.xlsx", before, 5)


def download_button(downloads: Path, name: str) -> FakeElement:
    button = element("AXPopUpButton", AXDescription="More Options")
    item = menu_item("Download")
    menu = element("AXMenu", item)

    def show() -> None:
        button.add(menu)

    def download() -> None:
        (downloads / name).write_bytes(CONTENT)
        menu.remove()

    button.actions["AXShowMenu"] = show
    item.actions["AXPress"] = download
    return button


def attachment_rig(tmp_path: Path, unread: bool = True) -> tuple[MailRig, AttachmentService, str]:
    downloads = tmp_path / "Downloads"
    downloads.mkdir()
    row = mail_row("Alice Example", "Project Atlas budget", "9:15 AM", "Numbers", unread=unread)
    gateway = MailGateway(rows=outlook_rows(row))
    button = download_button(downloads, "budget.xlsx")
    gateway.panes["row0"] = pane_for(
        PaneSpec(
            subject="Project Atlas budget",
            attachments=(
                attachment_group("budget.xlsx", "12.5 KB", button),
                attachment_group("notes.txt", "1 KB"),
            ),
        )
    )
    rig = mail_rig(gateway)
    driver = rig.rig.driver
    clock = FakeClock()
    service = AttachmentService(
        rig.core,
        gateway,
        driver.popups,
        PathDownloadsWatcher(downloads, clock, clock.sleep),
        tmp_path / "saved",
        driver.waiter,
    )
    key = rig.mail.conversations("Inbox", False, None, 0)[0].key
    return rig, service, key


def test_save_downloads_moves_and_restores_unread(tmp_path: Path) -> None:
    rig, service, key = attachment_rig(tmp_path)

    saved = service.save(key, "budget.xlsx", None, True)

    assert saved.to_record() == {
        "path": str(tmp_path / "saved" / "budget.xlsx"),
        "name": "budget.xlsx",
        "size_bytes": len(CONTENT),
        "notes": [],
    }
    assert not (tmp_path / "Downloads" / "budget.xlsx").exists()
    assert rig.marker.presses == 1
    assert rig.gateway.is_unread("row0")
    assert rig.gateway.selected == set()


def test_save_refuses_text_refs_before_touching_outlook(tmp_path: Path) -> None:
    rig, service, _ = attachment_rig(tmp_path)
    rig.gateway.calls.clear()

    with pytest.raises(TargetError, match="use the key"):
        service.save("budget", "budget.xlsx", None, True)
    assert not any(call.startswith("select:") for call in rig.gateway.calls)


def test_unknown_attachment_lists_the_names_and_restores_unread(tmp_path: Path) -> None:
    rig, service, key = attachment_rig(tmp_path)

    with pytest.raises(AttachmentError, match=r"budget\.xlsx, notes\.txt"):
        service.save(key, "missing.pdf", None, True)
    assert rig.gateway.is_unread("row0")
    assert rig.gateway.selected == set()


def test_attachment_without_a_menu_is_refused(tmp_path: Path) -> None:
    rig, service, key = attachment_rig(tmp_path, unread=False)

    with pytest.raises(AttachmentError, match="has no menu"):
        service.save(key, "notes.txt", None, True)
    assert rig.marker.presses == 0
    assert rig.gateway.selected == set()
