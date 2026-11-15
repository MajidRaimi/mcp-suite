from __future__ import annotations

from typing import Any, Literal

from office_mcp.features.persona import descriptions
from office_mcp.features.persona.guard import PersonaGuard
from office_mcp.features.persona.models import PersonaState
from office_mcp.features.persona.render import snapshot_words
from office_mcp.features.persona.store import PersonaStore
from office_mcp.shared.drafts import Channel, Draft, Review, Severity, Violation
from office_mcp.shared.errors import PersonaError, error_text
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import Effect, ToolGroup, ToolSpec
from office_mcp.shared.runner import Runner

Record = dict[str, Any]
ChannelName = Literal["email", "teams", "rsvp"]
UNAVAILABLE_RULE = "persona.unavailable"


class PersonaTools:
    def __init__(
        self, runner: Runner, store: PersonaStore, guard: PersonaGuard, owner: Owner
    ) -> None:
        self._runner = runner
        self._store = store
        self._guard = guard
        self._owner = owner

    def specs(self) -> list[ToolSpec]:
        return [
            ToolSpec(
                name="office_persona",
                function=self.office_persona,
                description=descriptions.office_persona(self._owner),
                group=ToolGroup.READ,
                effect=Effect.READS,
            ),
            ToolSpec(
                name="office_check_text",
                function=self.office_check_text,
                description=descriptions.office_check_text(self._owner),
                group=ToolGroup.READ,
                effect=Effect.READS,
            ),
        ]

    def office_persona(self) -> Record:
        return self._runner.unlocked(lambda: persona_record(self._store))

    def office_check_text(
        self,
        text: str,
        channel: ChannelName,
        subject: str = "",
        context: str = "",
        account: str = "",
    ) -> Record:
        draft = Draft(
            channel=Channel(channel),
            body=text,
            subject=subject,
            context=context,
            sender_account=account,
        )
        return self._runner.unlocked(lambda: self._check(draft).to_record())

    def _check(self, draft: Draft) -> Review:
        try:
            return self._guard.review(draft)
        except PersonaError as error:
            return unavailable_review(draft, error, self._store.version)


def persona_record(store: PersonaStore) -> Record:
    snapshot = store.snapshot()
    persona = snapshot.persona
    return {
        **snapshot.status.to_record(),
        "prose": "" if persona is None else persona.prose,
        "rules": {} if persona is None else persona.rules.to_record(),
        "rules_in_words": {channel.value: snapshot_words(snapshot, channel) for channel in Channel},
        "enforced": snapshot.status.state is not PersonaState.NONE,
    }


def unavailable_review(draft: Draft, error: PersonaError, version: str) -> Review:
    violation = Violation(
        rule=UNAVAILABLE_RULE,
        message=f"{error_text(error)}; every send is blocked until the persona file is fixed",
        severity=Severity.ERROR,
    )
    return Review(draft=draft, violations=(violation,), fixes=(), persona_version=version)
