from __future__ import annotations

from collections.abc import Callable
from typing import Annotated, Any, Literal

from pydantic import Field

from office_mcp.features.teams.writing import descriptions
from office_mcp.features.teams.writing.reactions import Reaction
from office_mcp.features.teams.writing.send import SendRequest
from office_mcp.features.teams.writing.service import TeamsWritingService
from office_mcp.features.teams.writing.targets import MAX_TEXT
from office_mcp.shared.drafts import Channel
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
    ToolSpec,
)
from office_mcp.shared.runner import SEND_BUDGET, Deadline, Lazy, Runner

Record = dict[str, Any]
ChatRef = Annotated[str, Field(min_length=1, max_length=512)]
MessageText = Annotated[str, Field(min_length=1, max_length=MAX_TEXT)]
MessageId = Annotated[str, Field(min_length=1, max_length=20, pattern=r"^\s*\d+\s*$")]
ReactionName = Literal["like", "heart", "laugh", "surprised"]


class TeamsWritingTools:
    def __init__(
        self,
        runner: Runner,
        writing: Lazy[TeamsWritingService],
        owner: Owner,
        rules_text: Callable[[Channel], str],
    ) -> None:
        self._runner = runner
        self._writing = writing
        self._owner = owner
        self._rules_text = rules_text

    def specs(self) -> list[ToolSpec]:
        owner = self._owner
        rules = self._rules_text(Channel.TEAMS)
        return [
            ToolSpec(
                "teams_send",
                self.teams_send,
                descriptions.teams_send(owner, rules),
                ToolGroup.SEND,
                Effect.SENDS,
            ),
            ToolSpec(
                "teams_reply",
                self.teams_reply,
                descriptions.teams_reply(owner, rules),
                ToolGroup.SEND,
                Effect.SENDS,
            ),
            ToolSpec(
                "teams_react",
                self.teams_react,
                descriptions.teams_react(owner),
                ToolGroup.SEND,
                Effect.SENDS,
            ),
            ToolSpec(
                "teams_delete_message",
                self.teams_delete_message,
                descriptions.teams_delete_message(owner),
                ToolGroup.SEND,
                Effect.SOFT_DELETES,
            ),
            ToolSpec(
                "teams_mark_unread",
                self.teams_mark_unread,
                EXPERIMENTAL_PREFIX + descriptions.teams_mark_unread(owner),
                ToolGroup.ORGANIZE,
                Effect.ORGANIZES,
                Stability.EXPERIMENTAL,
            ),
        ]

    def teams_send(
        self,
        chat: ChatRef,
        text: MessageText,
        group_ok: bool = False,
        allow_language_mismatch: bool = False,
        allow_duplicate: bool = False,
        dry_run: bool = False,
    ) -> Record:
        request = SendRequest(
            chat=chat,
            text=text,
            group_ok=group_ok,
            allow_language_mismatch=allow_language_mismatch,
            allow_duplicate=allow_duplicate,
            dry_run=dry_run,
        )
        return self._send(request)

    def teams_reply(
        self,
        chat: ChatRef,
        message_id: MessageId,
        text: MessageText,
        group_ok: bool = False,
        allow_language_mismatch: bool = False,
        allow_duplicate: bool = False,
        dry_run: bool = False,
    ) -> Record:
        request = SendRequest(
            chat=chat,
            text=text,
            group_ok=group_ok,
            allow_language_mismatch=allow_language_mismatch,
            allow_duplicate=allow_duplicate,
            dry_run=dry_run,
            reply_to=message_id,
        )
        return self._send(request)

    def teams_react(
        self,
        chat: ChatRef,
        message_id: MessageId,
        reaction: ReactionName,
        remove: bool = False,
    ) -> Record:
        wanted = Reaction(reaction)
        return self._runner.write(
            lambda deadline: (
                self._writing.get().react(chat, message_id, wanted, remove, deadline).to_record()
            )
        )

    def teams_delete_message(self, chat: ChatRef, message_id: MessageId) -> Record:
        return self._runner.write(
            lambda deadline: (
                self._writing.get().delete_message(chat, message_id, deadline).to_record()
            )
        )

    def teams_mark_unread(self, chat: ChatRef) -> Record:
        return self._runner.write(
            lambda deadline: self._writing.get().mark_unread(chat, deadline).to_record()
        )

    def _send(self, request: SendRequest) -> Record:
        def work(deadline: Deadline) -> Record:
            return self._writing.get().send(request, deadline).to_record()

        return self._runner.write(work, SEND_BUDGET)
