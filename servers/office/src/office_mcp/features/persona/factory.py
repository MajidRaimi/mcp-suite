from __future__ import annotations

from datetime import UTC, datetime

from office_mcp.features.persona.guard import PersonaGuard
from office_mcp.features.persona.store import PersonaStore
from office_mcp.shared.settings import Settings


def build_persona(settings: Settings) -> PersonaStore:
    return PersonaStore(
        path=settings.persona_file,
        explicit=settings.persona_explicit,
        required=settings.persona_required,
        clock=lambda: datetime.now(UTC),
    )


def build_persona_guard(store: PersonaStore) -> PersonaGuard:
    return PersonaGuard(store)
