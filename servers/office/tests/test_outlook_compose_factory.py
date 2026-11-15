from __future__ import annotations

from fakes.drivers import fake_runtime, outlook_rig
from fakes.guard import RecordingGuard
from fakes.p3_compose import FakeReader, inbox_key, latest_message

from office_mcp.features.outlook.compose.factory import build_compose, owner_address
from office_mcp.features.outlook.compose.service import ComposeService


def test_owner_address_prefers_the_configured_value() -> None:
    rig = outlook_rig()

    assert owner_address("alex@example.com", rig.app.app, rig.driver.windows)() == (
        "alex@example.com"
    )


def test_owner_address_reads_the_main_window_title() -> None:
    rig = outlook_rig()

    assert owner_address(None, rig.app.app, rig.driver.windows)() == "jane@example.com"


def test_owner_address_is_unknown_when_outlook_is_not_running() -> None:
    rig = outlook_rig()
    rig.app.processes.pids.clear()

    assert owner_address(None, rig.app.app, rig.driver.windows)() is None
    assert rig.app.processes.launched == []


def test_build_compose_wires_a_service() -> None:
    rig = outlook_rig()
    reader = FakeReader(latest_message(inbox_key()))

    service = build_compose(fake_runtime(), rig.driver, reader, RecordingGuard())

    assert isinstance(service, ComposeService)
