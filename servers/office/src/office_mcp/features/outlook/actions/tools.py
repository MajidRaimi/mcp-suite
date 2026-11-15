from __future__ import annotations

from collections.abc import Callable
from typing import Annotated, Any, Literal

from pydantic import Field

from office_mcp.features.outlook.actions import descriptions as text
from office_mcp.features.outlook.actions.flags import FlagState
from office_mcp.features.outlook.actions.service import MAX_REFS, ActionsService
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    Effect,
    Stability,
    ToolGroup,
    ToolSpec,
)
from office_mcp.shared.results import ActionResult
from office_mcp.shared.runner import BATCH_BUDGET, Deadline, Lazy, Runner

Record = dict[str, Any]
Keys = Annotated[list[str], Field(min_length=1, max_length=MAX_REFS, description=text.REFS_HELP)]
FolderHint = Annotated[str | None, Field(description=text.FOLDER_HELP)]
DestinationName = Annotated[
    str, Field(min_length=1, max_length=256, description=text.DESTINATION_HELP)
]
ReadFlag = Annotated[bool, Field(description=text.READ_HELP)]
FlagChoice = Annotated[
    Literal["today", "tomorrow", "this_week", "next_week", "no_date", "clear"],
    Field(description=text.STATE_HELP),
]
Work = Callable[[ActionsService, Deadline], ActionResult]


class ActionsTools:
    def __init__(self, runner: Runner, actions: Lazy[ActionsService], owner: Owner) -> None:
        self._runner = runner
        self._actions = actions
        self._owner = owner

    def outlook_archive(self, refs: Keys, folder: FolderHint = None) -> Record:
        return self._batch(lambda service, deadline: service.archive(refs, folder, deadline))

    def outlook_delete(self, refs: Keys, folder: FolderHint = None) -> Record:
        return self._batch(lambda service, deadline: service.delete(refs, folder, deadline))

    def outlook_move(
        self, refs: Keys, destination: DestinationName, folder: FolderHint = None
    ) -> Record:
        return self._batch(
            lambda service, deadline: service.move(refs, destination, folder, deadline)
        )

    def outlook_mark(self, refs: Keys, read: ReadFlag, folder: FolderHint = None) -> Record:
        return self._batch(lambda service, deadline: service.mark(refs, read, folder, deadline))

    def outlook_flag(
        self, refs: Keys, state: FlagChoice = "no_date", folder: FolderHint = None
    ) -> Record:
        chosen = FlagState(state)
        return self._batch(lambda service, deadline: service.flag(refs, chosen, folder, deadline))

    def specs(self) -> list[ToolSpec]:
        owner = self._owner
        return [
            organize(self.outlook_archive, text.outlook_archive(owner), Effect.ORGANIZES),
            organize(self.outlook_delete, text.outlook_delete(owner), Effect.SOFT_DELETES),
            organize(self.outlook_move, text.outlook_move(owner), Effect.ORGANIZES),
            organize(self.outlook_mark, text.outlook_mark(owner), Effect.ORGANIZES),
            organize(self.outlook_flag, text.outlook_flag(owner), Effect.ORGANIZES),
        ]

    def _batch(self, work: Work) -> Record:
        return self._runner.write(
            lambda deadline: work(self._actions.get(), deadline).to_record(), BATCH_BUDGET
        )


def organize(
    function: Callable[..., object],
    description: str,
    effect: Effect,
    stability: Stability = Stability.STABLE,
) -> ToolSpec:
    return ToolSpec(
        name=function.__name__,
        function=function,
        description=description,
        group=ToolGroup.ORGANIZE,
        effect=effect,
        stability=stability,
    )
