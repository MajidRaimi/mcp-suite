from __future__ import annotations

import itertools

import pytest

from office_mcp.shared.registry import (
    Effect,
    Stability,
    ToolGroup,
    ToolPolicy,
    ToolSpec,
    annotations_for,
    enabled_specs,
    unknown_names,
)
from office_mcp.shared.settings import Settings


def noop() -> None:
    return None


def spec(
    name: str,
    group: ToolGroup,
    stability: Stability = Stability.STABLE,
    effect: Effect = Effect.READS,
) -> ToolSpec:
    return ToolSpec(name, noop, f"{name} tool", group, effect, stability)


SPECS = [
    spec("read_stable", ToolGroup.READ),
    spec("read_exp", ToolGroup.READ, Stability.EXPERIMENTAL),
    spec("organize_stable", ToolGroup.ORGANIZE, effect=Effect.ORGANIZES),
    spec("organize_exp", ToolGroup.ORGANIZE, Stability.EXPERIMENTAL, Effect.ORGANIZES),
    spec("send_stable", ToolGroup.SEND, effect=Effect.SENDS),
    spec("send_exp", ToolGroup.SEND, Stability.EXPERIMENTAL, Effect.SENDS),
]
GROUP_SETS = [
    frozenset(combo)
    for size in range(1, 4)
    for combo in itertools.combinations(list(ToolGroup), size)
]


def expected(policy: ToolPolicy) -> set[str]:
    names = set()
    for item in SPECS:
        experimental_ok = item.stability is Stability.STABLE or (
            policy.all_experimental or item.name in policy.experimental
        )
        if item.group in policy.groups and item.name not in policy.disabled and experimental_ok:
            names.add(item.name)
    return names


@pytest.mark.parametrize("groups", GROUP_SETS)
@pytest.mark.parametrize("disabled", [frozenset(), frozenset({"read_stable", "send_exp"})])
@pytest.mark.parametrize(
    ("experimental", "all_experimental"),
    [(frozenset(), False), (frozenset({"send_exp", "read_exp"}), False), (frozenset(), True)],
)
def test_enabled_specs_for_every_combination(
    groups: frozenset[ToolGroup],
    disabled: frozenset[str],
    experimental: frozenset[str],
    all_experimental: bool,
) -> None:
    policy = ToolPolicy(groups, disabled, experimental, all_experimental)

    names = {item.name for item in enabled_specs(SPECS, policy)}

    assert names == expected(policy)


def test_read_only_hides_every_write_and_experimental_tool() -> None:
    policy = Settings.from_env(
        {
            "OFFICE_MCP_READ_ONLY": "true",
            "OFFICE_MCP_TOOLS": "all",
            "OFFICE_MCP_EXPERIMENTAL": "all",
            "OFFICE_MCP_DISABLED_TOOLS": "read_stable",
        }
    ).tool_policy

    assert [item.name for item in enabled_specs(SPECS, policy)] == []
    assert policy.groups == frozenset({ToolGroup.READ})


def test_unknown_names_lists_disabled_and_experimental_names_missing_from_specs() -> None:
    policy = ToolPolicy(
        frozenset({ToolGroup.READ}),
        frozenset({"read_stable", "outlook_purge"}),
        frozenset({"send_exp", "teams_call"}),
        False,
    )

    assert unknown_names(SPECS, policy) == ("outlook_purge", "teams_call")


@pytest.mark.parametrize(
    ("effect", "hints"),
    [
        (Effect.READS, (True, False, True, False)),
        (Effect.MARKS_READ, (False, False, False, False)),
        (Effect.ORGANIZES, (False, False, True, False)),
        (Effect.SOFT_DELETES, (False, True, True, False)),
        (Effect.WRITES_LOCAL, (False, False, False, False)),
        (Effect.SENDS, (False, False, False, True)),
    ],
)
def test_annotations_table(effect: Effect, hints: tuple[bool, bool, bool, bool]) -> None:
    annotations = annotations_for(effect)

    assert (
        annotations.read_only_hint,
        annotations.destructive_hint,
        annotations.idempotent_hint,
        annotations.open_world_hint,
    ) == hints
