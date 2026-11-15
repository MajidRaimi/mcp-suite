from __future__ import annotations

from dataclasses import dataclass, replace

from office_mcp.features.persona.models import ChannelRules, Rules
from office_mcp.features.persona.rules import (
    characters,
    emoji,
    emoticons,
    language,
    length,
    patterns,
    phrases,
    signoff,
)
from office_mcp.features.persona.rules.common import fields
from office_mcp.shared.drafts import (
    Channel,
    Draft,
    Review,
    Severity,
    Violation,
    canonical_text,
)


@dataclass(frozen=True)
class PersonaPolicy:
    rules: Rules
    version: str

    def review(self, draft: Draft) -> Review:
        body = canonical_text(draft.body)
        subject = canonical_text(draft.subject)
        channel_rules = self.rules.channel(draft.channel)
        outcome = self._sign_off(draft, body, channel_rules)
        final = replace(draft, body=outcome.body, subject=subject)
        spoken = self._language(final)
        violations = (
            *outcome.violations,
            *self._content(final),
            *(item for item in spoken if item.severity is Severity.ERROR),
            *([] if channel_rules is None else length.check(final.body, channel_rules)),
            *(item for item in spoken if item.severity is Severity.WARNING),
        )
        return Review(
            draft=final,
            violations=violations,
            fixes=outcome.fixes,
            persona_version=self.version,
        )

    def _sign_off(
        self, draft: Draft, body: str, rules: ChannelRules | None
    ) -> signoff.SignOffOutcome:
        if rules is None:
            return signoff.SignOffOutcome(body=body)
        if draft.channel is Channel.EMAIL and not rules.applies_to(draft.sender_account):
            return signoff.SignOffOutcome(body=body)
        written = language.language_of(language.language_sample(body))
        return signoff.apply(body, rules, written)

    def _content(self, draft: Draft) -> list[Violation]:
        forbid = self.rules.forbid
        targets = fields(draft.body, draft.subject)
        teams_emoticons = forbid.emoji and draft.channel is Channel.TEAMS
        return [
            *characters.check(targets, forbid.forbidden_characters()),
            *(emoji.check(targets, forbid.emoji_allow) if forbid.emoji else []),
            *(emoticons.check(targets) if teams_emoticons else []),
            *phrases.check(targets, forbid.phrases),
            *patterns.check(targets, forbid.patterns),
        ]

    def _language(self, draft: Draft) -> list[Violation]:
        if self.rules.language is None:
            return []
        return language.check(draft, draft.body, self.rules.language)
