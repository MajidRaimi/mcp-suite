from __future__ import annotations

from dataclasses import dataclass

DEFAULT_OWNER = "the user"


@dataclass(frozen=True)
class Owner:
    name: str | None = None
    email: str | None = None

    @property
    def label(self) -> str:
        return self.name or DEFAULT_OWNER

    @property
    def possessive(self) -> str:
        label = self.label
        return f"{label}'" if label.endswith("s") and self.name else f"{label}'s"
