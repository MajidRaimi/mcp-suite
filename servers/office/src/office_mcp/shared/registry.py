from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass
from enum import StrEnum

from mcp.types import ToolAnnotations


class ToolGroup(StrEnum):
    READ = "read"
    ORGANIZE = "organize"
    SEND = "send"


class Effect(StrEnum):
    READS = "reads"
    MARKS_READ = "marks_read"
    ORGANIZES = "organizes"
    SOFT_DELETES = "soft_deletes"
    WRITES_LOCAL = "writes_local"
    SENDS = "sends"


class Stability(StrEnum):
    STABLE = "stable"
    EXPERIMENTAL = "experimental"


EXPERIMENTAL_PREFIX = "Experimental: "


@dataclass(frozen=True)
class ToolSpec:
    name: str
    function: Callable[..., object]
    description: str
    group: ToolGroup
    effect: Effect
    stability: Stability = Stability.STABLE


@dataclass(frozen=True)
class ToolPolicy:
    groups: frozenset[ToolGroup]
    disabled: frozenset[str]
    experimental: frozenset[str]
    all_experimental: bool

    def allows(self, spec: ToolSpec) -> bool:
        if spec.group not in self.groups or spec.name in self.disabled:
            return False
        if spec.stability is Stability.STABLE:
            return True
        return self.all_experimental or spec.name in self.experimental


@dataclass(frozen=True)
class Hints:
    read_only: bool
    destructive: bool
    idempotent: bool
    open_world: bool


EFFECT_HINTS: dict[Effect, Hints] = {
    Effect.READS: Hints(read_only=True, destructive=False, idempotent=True, open_world=False),
    Effect.MARKS_READ: Hints(
        read_only=False, destructive=False, idempotent=False, open_world=False
    ),
    Effect.ORGANIZES: Hints(read_only=False, destructive=False, idempotent=True, open_world=False),
    Effect.SOFT_DELETES: Hints(
        read_only=False, destructive=True, idempotent=True, open_world=False
    ),
    Effect.WRITES_LOCAL: Hints(
        read_only=False, destructive=False, idempotent=False, open_world=False
    ),
    Effect.SENDS: Hints(read_only=False, destructive=False, idempotent=False, open_world=True),
}


def annotations_for(effect: Effect) -> ToolAnnotations:
    hints = EFFECT_HINTS[effect]
    return ToolAnnotations(
        read_only_hint=hints.read_only,
        destructive_hint=hints.destructive,
        idempotent_hint=hints.idempotent,
        open_world_hint=hints.open_world,
    )


def enabled_specs(specs: Iterable[ToolSpec], policy: ToolPolicy) -> list[ToolSpec]:
    return [spec for spec in specs if policy.allows(spec)]


def unknown_names(specs: Iterable[ToolSpec], policy: ToolPolicy) -> tuple[str, ...]:
    known = {spec.name for spec in specs}
    return tuple(sorted((policy.disabled | policy.experimental) - known))
