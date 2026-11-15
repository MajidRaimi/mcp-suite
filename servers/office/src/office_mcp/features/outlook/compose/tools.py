from __future__ import annotations

from collections.abc import Callable
from typing import Annotated, Any

from pydantic import Field

from office_mcp.features.outlook.compose import descriptions as text
from office_mcp.features.outlook.compose.recipients import MAX_PER_FIELD
from office_mcp.features.outlook.compose.sensitivity import MAX_LABEL, MAX_LEVELS
from office_mcp.features.outlook.compose.service import (
    MAX_BODY,
    MAX_SUBJECT,
    ComposeService,
    ForwardMail,
    NewMail,
    ReplyMail,
)
from office_mcp.shared.drafts import Channel
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import (
    EXPERIMENTAL_PREFIX,
    Effect,
    Stability,
    ToolGroup,
    ToolSpec,
)
from office_mcp.shared.results import SendResult
from office_mcp.shared.runner import SEND_BUDGET, WRITE_BUDGET, Deadline, Lazy, Runner

Record = dict[str, Any]
Recipients = Annotated[list[str], Field(min_length=1, max_length=MAX_PER_FIELD)]
ExtraRecipients = Annotated[list[str] | None, Field(max_length=MAX_PER_FIELD)]
Subject = Annotated[str, Field(min_length=1, max_length=MAX_SUBJECT)]
Body = Annotated[str, Field(min_length=1, max_length=MAX_BODY)]
OptionalBody = Annotated[str, Field(max_length=MAX_BODY)]
Key = Annotated[str, Field(min_length=12, max_length=12)]
FolderHint = Annotated[str | None, Field(max_length=256)]
SensitivityLabel = Annotated[str, Field(min_length=1, max_length=MAX_LABEL)]
Sensitivity = Annotated[list[SensitivityLabel] | None, Field(min_length=1, max_length=MAX_LEVELS)]
Work = Callable[[ComposeService, Deadline], SendResult]


class ComposeTools:
    def __init__(
        self,
        runner: Runner,
        compose: Lazy[ComposeService],
        owner: Owner,
        rules_text: Callable[[Channel], str],
    ) -> None:
        self._runner = runner
        self._compose = compose
        self._owner = owner
        self._rules_text = rules_text

    def outlook_send(
        self,
        to: Recipients,
        subject: Subject,
        body: Body,
        cc: ExtraRecipients = None,
        bcc: ExtraRecipients = None,
        allow_language_mismatch: bool = False,
        allow_duplicate: bool = False,
        dry_run: bool = False,
        sensitivity: Sensitivity = None,
    ) -> Record:
        mail = NewMail(
            to=tuple(to),
            subject=subject,
            body=body,
            cc=tuple(cc or ()),
            bcc=tuple(bcc or ()),
            allow_language_mismatch=allow_language_mismatch,
            allow_duplicate=allow_duplicate,
            dry_run=dry_run,
            sensitivity=tuple(sensitivity or ()),
        )
        return self._send(lambda service, deadline: service.send(mail, deadline))

    def outlook_reply(
        self,
        ref: Key,
        body: Body,
        reply_all: bool = False,
        cc: ExtraRecipients = None,
        bcc: ExtraRecipients = None,
        folder: FolderHint = None,
        allow_language_mismatch: bool = False,
        allow_duplicate: bool = False,
        dry_run: bool = False,
        sensitivity: Sensitivity = None,
    ) -> Record:
        mail = ReplyMail(
            key=ref,
            body=body,
            reply_all=reply_all,
            cc=tuple(cc or ()),
            bcc=tuple(bcc or ()),
            folder=folder,
            allow_language_mismatch=allow_language_mismatch,
            allow_duplicate=allow_duplicate,
            dry_run=dry_run,
            sensitivity=tuple(sensitivity or ()),
        )
        return self._send(lambda service, deadline: service.reply(mail, deadline))

    def outlook_forward(
        self,
        ref: Key,
        to: Recipients,
        body: OptionalBody = "",
        cc: ExtraRecipients = None,
        bcc: ExtraRecipients = None,
        folder: FolderHint = None,
        allow_duplicate: bool = False,
        dry_run: bool = False,
        sensitivity: Sensitivity = None,
    ) -> Record:
        mail = ForwardMail(
            key=ref,
            to=tuple(to),
            body=body,
            cc=tuple(cc or ()),
            bcc=tuple(bcc or ()),
            folder=folder,
            allow_duplicate=allow_duplicate,
            dry_run=dry_run,
            sensitivity=tuple(sensitivity or ()),
        )
        return self._send(lambda service, deadline: service.forward(mail, deadline))

    def outlook_undo_send(self) -> Record:
        return self._runner.write(
            lambda _deadline: self._compose.get().undo_send().to_record(), WRITE_BUDGET
        )

    def specs(self) -> list[ToolSpec]:
        owner = self._owner
        rules = self._rules_text(Channel.EMAIL)
        return [
            send_spec(self.outlook_send, text.outlook_send(owner, rules)),
            send_spec(self.outlook_reply, text.outlook_reply(owner, rules)),
            send_spec(self.outlook_forward, text.outlook_forward(owner, rules)),
            send_spec(
                self.outlook_undo_send,
                EXPERIMENTAL_PREFIX + text.outlook_undo_send(owner),
                Stability.EXPERIMENTAL,
            ),
        ]

    def _send(self, work: Work) -> Record:
        return self._runner.write(
            lambda deadline: work(self._compose.get(), deadline).to_record(), SEND_BUDGET
        )


def send_spec(
    function: Callable[..., object],
    description: str,
    stability: Stability = Stability.STABLE,
) -> ToolSpec:
    return ToolSpec(
        name=function.__name__,
        function=function,
        description=description,
        group=ToolGroup.SEND,
        effect=Effect.SENDS,
        stability=stability,
    )
