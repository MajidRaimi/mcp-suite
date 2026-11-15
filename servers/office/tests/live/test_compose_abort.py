from __future__ import annotations

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from live.conftest import LiveRig
from office_mcp.features.outlook.compose.service import NewMail
from office_mcp.shared.runner import SEND_BUDGET

pytestmark = pytest.mark.live


def test_compose_abort_discards_the_draft_without_sending(live: LiveRig) -> None:
    compose = live.container.compose.get()
    mail = NewMail(
        to=(live.targets.account,),
        subject=live.subject("abort"),
        body="This compose is closed before Send on purpose.",
    )

    with pytest.raises(ToolError, match="copy is in Deleted Items"):
        live.container.runner.write(
            lambda deadline: compose.send(mail, deadline, fail_at="after_body"), SEND_BUDGET
        )
