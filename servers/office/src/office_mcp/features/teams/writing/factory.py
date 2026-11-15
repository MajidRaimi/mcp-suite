from __future__ import annotations

from collections.abc import Callable

from office_mcp.features.teams.driver import TeamsDriver
from office_mcp.features.teams.writing.compose import ComposeBox
from office_mcp.features.teams.writing.context import WriteContext
from office_mcp.features.teams.writing.deletion import DeletionService
from office_mcp.features.teams.writing.guard import TeamsSendGuard
from office_mcp.features.teams.writing.reactions import ReactionService
from office_mcp.features.teams.writing.reply import Quoter
from office_mcp.features.teams.writing.send import TeamsSender
from office_mcp.features.teams.writing.service import TeamsWritingService
from office_mcp.features.teams.writing.targets import ChatTargets
from office_mcp.features.teams.writing.unread import UnreadService
from office_mcp.shared.drafts import TextGuard
from office_mcp.shared.recipients import RecipientGuard
from office_mcp.shared.runtime import Runtime


def writing_context(teams: TeamsDriver, recipients: RecipientGuard) -> WriteContext:
    targets = ChatTargets(teams.chats, teams.opener, teams.gateway, teams.waiter, teams.zone)
    return WriteContext(
        targets=targets,
        gateway=teams.gateway,
        scopes=teams.scope,
        waiter=teams.waiter,
        recipients=recipients,
    )


def assemble_writing(
    context: WriteContext,
    box: ComposeBox,
    guard: TextGuard,
    clock: Callable[[], float],
    max_members: int,
) -> TeamsWritingService:
    sender = TeamsSender(
        context=context,
        box=box,
        quoter=Quoter(context, box),
        guard=guard,
        duplicates=TeamsSendGuard(clock),
        max_members=max_members,
    )
    return TeamsWritingService(
        sender=sender,
        reactions=ReactionService(context),
        deletion=DeletionService(context),
        unread=UnreadService(context),
    )


def build_teams_writing(
    runtime: Runtime, teams: TeamsDriver, guard: TextGuard
) -> TeamsWritingService:
    return assemble_writing(
        context=writing_context(teams, runtime.recipients),
        box=ComposeBox(teams.gateway, teams.text),
        guard=guard,
        clock=runtime.waiter.clock,
        max_members=runtime.settings.teams_max_members,
    )
