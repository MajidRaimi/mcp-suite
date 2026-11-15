from __future__ import annotations

from office_mcp.features.teams.writing.deletion import DeletionService
from office_mcp.features.teams.writing.reactions import Reaction, ReactionResult, ReactionService
from office_mcp.features.teams.writing.send import SendRequest, TeamsSender
from office_mcp.features.teams.writing.unread import UnreadService
from office_mcp.shared.results import ActionResult, SendResult
from office_mcp.shared.runner import Deadline


class TeamsWritingService:
    def __init__(
        self,
        sender: TeamsSender,
        reactions: ReactionService,
        deletion: DeletionService,
        unread: UnreadService,
    ) -> None:
        self._sender = sender
        self._reactions = reactions
        self._deletion = deletion
        self._unread = unread

    def send(self, request: SendRequest, deadline: Deadline) -> SendResult:
        return self._sender.send(request, deadline)

    def react(
        self, chat: str, message_id: str, reaction: Reaction, remove: bool, deadline: Deadline
    ) -> ReactionResult:
        return self._reactions.react(chat, message_id, reaction, remove, deadline)

    def delete_message(self, chat: str, message_id: str, deadline: Deadline) -> ActionResult:
        return self._deletion.delete(chat, message_id, deadline)

    def mark_unread(self, chat: str, deadline: Deadline) -> ActionResult:
        return self._unread.mark_unread(chat, deadline)
