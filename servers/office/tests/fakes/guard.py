from __future__ import annotations

from office_mcp.shared.drafts import Draft, Review, require_written
from office_mcp.shared.errors import OfficeError, PersonaViolationError


class RecordingGuard:
    def __init__(self, scripted: list[Review | OfficeError] | None = None) -> None:
        self.scripted = list(scripted or [])
        self.reviewed: list[Draft] = []
        self.verified: list[tuple[Draft, str, str]] = []
        self.persona_version = "test"

    def review(self, draft: Draft) -> Review:
        self.reviewed.append(draft)
        if not self.scripted:
            return Review(
                draft=draft, violations=(), fixes=(), persona_version=self.persona_version
            )
        scripted = self.scripted.pop(0)
        if isinstance(scripted, OfficeError):
            raise scripted
        return scripted

    def require_clean(self, draft: Draft) -> Review:
        review = self.review(draft)
        if not review.ok:
            raise PersonaViolationError(review.violations)
        return review

    def verify_written(self, intended: Draft, body_read: str, subject_read: str) -> None:
        self.verified.append((intended, body_read, subject_read))
        require_written(intended, body_read, subject_read)
