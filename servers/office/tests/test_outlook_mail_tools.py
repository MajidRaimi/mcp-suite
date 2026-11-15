from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest
from fakes.ax_backend import FakeBackend, FakeProcesses
from fakes.clock import FakeClock
from fakes.drivers import fake_runtime
from fakes.lock import MemoryUiLock
from fakes.p1_mail import MailGateway, MailRig, PaneSpec, mail_rig, mail_row, reading_pane
from fakes.trees import folder_row, outlook_rows
from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.features.outlook.driver import build_outlook_driver
from office_mcp.features.outlook.mail.factory import build_mail
from office_mcp.features.outlook.mail.service import MailService
from office_mcp.features.outlook.mail.tools import MailTools
from office_mcp.features.outlook.rows import RefRegistry
from office_mcp.shared.errors import ToolDisabledError
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
    annotations_for,
)
from office_mcp.shared.runner import Lazy, Runner
from office_mcp.shared.settings import Settings

ROW = mail_row("Alice Example", "Project Atlas budget", "9:15 AM", "Numbers attached")
READ_TOOLS = [
    "outlook_folders",
    "outlook_conversations",
    "outlook_read",
    "outlook_search",
    "outlook_people",
    "outlook_needs_reply",
    "outlook_follow_ups",
]


def settings(download_dir: Path | None = None) -> Settings:
    base = Settings.from_env({"OFFICE_MCP_TIMEZONE": "UTC"})
    return replace(base, download_dir=download_dir)


def tools_for(rig: MailRig, current: Settings | None = None) -> MailTools:
    runner = Runner(MemoryUiLock(), FakeClock(), 20, 10)
    mail: Lazy[MailService] = Lazy(lambda: rig.mail)
    return MailTools(runner, mail, Owner("Jane Doe"), current or settings())


def rig_with_mail() -> MailRig:
    gateway = MailGateway(
        rows=outlook_rows(ROW),
        folders=[
            folder_row("jane@example.com", level=0),
            folder_row("Inbox; 1 unread message", handle="inbox"),
            folder_row("Sent Items", handle="sent"),
        ],
    )
    gateway.panes["row0"] = reading_pane(PaneSpec(subject="Project Atlas budget"))
    return mail_rig(gateway)


def test_default_specs_are_the_seven_read_tools() -> None:
    specs = tools_for(rig_with_mail()).specs()

    assert [spec.name for spec in specs] == READ_TOOLS
    assert {spec.group for spec in specs} == {ToolGroup.READ}
    assert {spec.stability for spec in specs} == {Stability.STABLE}
    effects = {spec.name: spec.effect for spec in specs}
    assert effects["outlook_read"] is Effect.MARKS_READ
    assert {effect for name, effect in effects.items() if name != "outlook_read"} == {Effect.READS}
    assert all(spec.function.__name__ == spec.name for spec in specs)


def test_save_attachment_appears_only_with_a_download_dir(tmp_path: Path) -> None:
    specs = tools_for(rig_with_mail(), settings(tmp_path)).specs()

    extra = specs[-1]
    assert [spec.name for spec in specs] == [*READ_TOOLS, "outlook_save_attachment"]
    assert extra.group is ToolGroup.ORGANIZE
    assert extra.effect is Effect.WRITES_LOCAL
    assert extra.stability is Stability.EXPERIMENTAL
    assert extra.description.startswith(EXPERIMENTAL_PREFIX)
    assert annotations_for(extra.effect).read_only_hint is False


def test_descriptions_have_no_long_dashes() -> None:
    for spec in tools_for(rig_with_mail(), settings(Path("/tmp/x"))).specs():
        assert chr(0x2013) not in spec.description
        assert chr(0x2014) not in spec.description


def test_tools_return_documented_records() -> None:
    rig = rig_with_mail()
    tools = tools_for(rig)

    folders = tools.outlook_folders()
    conversations = tools.outlook_conversations(view="both")
    email = tools.outlook_read(conversations[0]["key"])
    needs = tools.outlook_needs_reply()
    follow = tools.outlook_follow_ups()
    found = tools.outlook_search("atlas")
    people = tools.outlook_people("al")

    assert folders[1:] == [
        {
            "name": "Inbox",
            "path": "jane@example.com/Inbox",
            "level": 1,
            "unread": 1,
            "total": None,
            "unseen": None,
            "favorite": False,
        },
        {
            "name": "Sent Items",
            "path": "jane@example.com/Sent Items",
            "level": 1,
            "unread": 0,
            "total": None,
            "unseen": None,
            "favorite": False,
        },
    ]
    assert {"key", "folder", "view", "flagged", "replied", "forwarded"} <= set(conversations[0])
    assert {"messages", "attachments", "conversation", "body"} <= set(email)
    assert needs[0]["reason"] == "read, no reply yet"
    assert follow == []
    assert set(found) == {"query", "scope", "count", "results"}
    assert people == [
        {"name": "Alice Example", "email": "alice@example.com", "source": "recent_recipient"}
    ]


def test_save_attachment_without_a_service_is_a_tool_error(tmp_path: Path) -> None:
    tools = tools_for(rig_with_mail(), settings(tmp_path))

    with pytest.raises(ToolError, match="OFFICE_MCP_DOWNLOAD_DIR"):
        tools.outlook_save_attachment("0123456789ab", "budget.xlsx")


def test_build_mail_wires_attachments_only_with_a_download_dir(tmp_path: Path) -> None:
    runtime = fake_runtime(FakeBackend(), FakeProcesses())
    driver = build_outlook_driver(runtime, RefRegistry())
    with_dir = replace(runtime, settings=replace(runtime.settings, download_dir=tmp_path))

    without = build_mail(runtime, driver)
    configured = build_mail(with_dir, driver)

    assert isinstance(without, MailService)
    assert isinstance(configured, MailService)
    with pytest.raises(ToolDisabledError, match="OFFICE_MCP_DOWNLOAD_DIR"):
        without.save_attachment("0123456789ab", "a.txt", None, True)
