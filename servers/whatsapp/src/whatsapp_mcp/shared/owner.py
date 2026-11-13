from __future__ import annotations

from dataclasses import dataclass

GENERIC_OWNER = "the user"


@dataclass(frozen=True)
class Owner:
    name: str | None = None

    @property
    def label(self) -> str:
        return self.name or GENERIC_OWNER

    @property
    def possessive(self) -> str:
        return f"{self.label}'s"
