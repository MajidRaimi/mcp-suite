from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path

import pytest

from live.guard import (
    RUN_VARIABLE,
    LiveRefusal,
    SelfTargets,
    live_enabled,
    run_subject,
    self_targets,
)
from office_mcp.container import Container, build_container
from office_mcp.shared.settings import Settings
from office_mcp.tools import OfficeTools

LIVE_DIRECTORY = Path(__file__).resolve().parent


@dataclass(frozen=True)
class LiveRig:
    container: Container
    targets: SelfTargets
    run_id: str

    def subject(self, label: str) -> str:
        return run_subject(self.run_id, label)


def pytest_collection_modifyitems(items: list[pytest.Item]) -> None:
    for item in items:
        if LIVE_DIRECTORY in Path(str(item.path)).resolve().parents:
            item.add_marker(pytest.mark.live)


def run_id() -> str:
    return os.environ.get(RUN_VARIABLE, "").strip() or datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")


@pytest.fixture(scope="session")
def live() -> LiveRig:
    if not live_enabled(os.environ):
        pytest.skip("live tests drive the real apps; set OFFICE_MCP_LIVE=1 on the Mac")
    container = build_container(Settings.from_env())
    status = OfficeTools(container).office_status()
    try:
        targets = self_targets(status, os.environ)
    except LiveRefusal as refusal:
        pytest.exit(f"live run refused: {refusal}", returncode=2)
    return LiveRig(container=container, targets=targets, run_id=run_id())
