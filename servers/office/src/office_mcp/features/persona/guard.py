from __future__ import annotations

from dataclasses import replace

from office_mcp.features.persona.policy import PersonaPolicy
from office_mcp.features.persona.store import PersonaStore
from office_mcp.shared.drafts import Draft, Review, require_written
from office_mcp.shared.errors import PersonaViolationError


class PersonaGuard:
    def __init__(self, store: PersonaStore) -> None:
        self._store = store

    @property
    def persona_version(self) -> str:
        return self._store.version

    def review(self, draft: Draft) -> Review:
        persona = self._store.current()
        return PersonaPolicy(rules=persona.rules, version=persona.version).review(draft)

    def require_clean(self, draft: Draft) -> Review:
        review = self.review(draft)
        if not review.ok:
            raise PersonaViolationError(review.violations)
        return review

    def verify_written(self, intended: Draft, body_read: str, subject_read: str) -> None:
        require_written(intended, body_read, subject_read)
        written = replace(intended, body=body_read, subject=subject_read)
        self.require_clean(written)
