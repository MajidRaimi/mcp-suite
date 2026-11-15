from __future__ import annotations

import time
from dataclasses import dataclass
from functools import partial

from office_mcp.features.calendar.factory import build_calendar
from office_mcp.features.calendar.service import CalendarService
from office_mcp.features.calendar.tools import CalendarTools
from office_mcp.features.outlook.actions.factory import build_actions
from office_mcp.features.outlook.actions.service import ActionsService
from office_mcp.features.outlook.actions.tools import ActionsTools
from office_mcp.features.outlook.compose.factory import build_compose
from office_mcp.features.outlook.compose.service import ComposeService
from office_mcp.features.outlook.compose.tools import ComposeTools
from office_mcp.features.outlook.driver import OutlookDriver, build_outlook_driver
from office_mcp.features.outlook.mail.factory import build_mail
from office_mcp.features.outlook.mail.service import MailService
from office_mcp.features.outlook.mail.tools import MailTools
from office_mcp.features.outlook.rows import RefRegistry
from office_mcp.features.persona.factory import build_persona, build_persona_guard
from office_mcp.features.persona.guard import PersonaGuard
from office_mcp.features.persona.render import send_rules_text
from office_mcp.features.persona.store import PersonaStore
from office_mcp.features.persona.tools import PersonaTools
from office_mcp.features.teams.driver import TeamsDriver, build_teams_driver
from office_mcp.features.teams.reading.factory import build_teams_reading
from office_mcp.features.teams.reading.service import TeamsReadingService
from office_mcp.features.teams.reading.tools import TeamsReadingTools
from office_mcp.features.teams.writing.factory import build_teams_writing
from office_mcp.features.teams.writing.service import TeamsWritingService
from office_mcp.features.teams.writing.tools import TeamsWritingTools
from office_mcp.shared.errors import SettingsError
from office_mcp.shared.owner import Owner
from office_mcp.shared.registry import ToolSpec, enabled_specs, unknown_names
from office_mcp.shared.runner import Lazy, ProcessUiLock, Runner
from office_mcp.shared.runtime import Runtime, build_runtime
from office_mcp.shared.settings import Settings
from office_mcp.tools import OfficeTools

CONDITIONAL_TOOLS = frozenset({"outlook_save_attachment"})


@dataclass(frozen=True)
class Container:
    settings: Settings
    owner: Owner
    runner: Runner
    runtime: Lazy[Runtime]
    registry: RefRegistry
    outlook: Lazy[OutlookDriver]
    teams: Lazy[TeamsDriver]
    persona: PersonaStore
    persona_guard: PersonaGuard
    mail: Lazy[MailService]
    actions: Lazy[ActionsService]
    compose: Lazy[ComposeService]
    calendar: Lazy[CalendarService]
    teams_reading: Lazy[TeamsReadingService]
    teams_writing: Lazy[TeamsWritingService]


def build_runner(settings: Settings) -> Runner:
    lock = ProcessUiLock(settings.lock_file, time.monotonic, time.sleep)
    return Runner(lock, time.monotonic, settings.lock_timeout, settings.stale_after)


def build_container(settings: Settings) -> Container:
    runtime = Lazy(lambda: build_runtime(settings))
    registry = RefRegistry()
    outlook = Lazy(lambda: build_outlook_driver(runtime.get(), registry))
    teams = Lazy(lambda: build_teams_driver(runtime.get()))
    persona = build_persona(settings)
    persona_guard = build_persona_guard(persona)
    mail = Lazy(lambda: build_mail(runtime.get(), outlook.get()))
    return Container(
        settings=settings,
        owner=Owner(settings.owner_name, settings.owner_email),
        runner=build_runner(settings),
        runtime=runtime,
        registry=registry,
        outlook=outlook,
        teams=teams,
        persona=persona,
        persona_guard=persona_guard,
        mail=mail,
        actions=Lazy(lambda: build_actions(runtime.get(), outlook.get())),
        compose=Lazy(
            lambda: build_compose(runtime.get(), outlook.get(), mail.get(), persona_guard)
        ),
        calendar=Lazy(lambda: build_calendar(runtime.get(), teams.get(), persona_guard)),
        teams_reading=Lazy(lambda: build_teams_reading(runtime.get(), teams.get())),
        teams_writing=Lazy(lambda: build_teams_writing(runtime.get(), teams.get(), persona_guard)),
    )


def all_specs(container: Container) -> list[ToolSpec]:
    runner = container.runner
    owner = container.owner
    rules = partial(send_rules_text, container.persona)
    return [
        *OfficeTools(container).specs(),
        *PersonaTools(runner, container.persona, container.persona_guard, owner).specs(),
        *MailTools(runner, container.mail, owner, container.settings).specs(),
        *ActionsTools(runner, container.actions, owner).specs(),
        *ComposeTools(runner, container.compose, owner, rules).specs(),
        *CalendarTools(runner, container.calendar, owner, rules).specs(),
        *TeamsReadingTools(runner, container.teams_reading, owner).specs(),
        *TeamsWritingTools(runner, container.teams_writing, owner, rules).specs(),
    ]


def require_known_names(specs: list[ToolSpec], container: Container) -> None:
    unknown = [
        name
        for name in unknown_names(specs, container.settings.tool_policy)
        if name not in CONDITIONAL_TOOLS
    ]
    if unknown:
        raise SettingsError(
            "OFFICE_MCP_DISABLED_TOOLS or OFFICE_MCP_EXPERIMENTAL names unknown tools: "
            + ", ".join(unknown)
        )


def exposed_specs(container: Container) -> list[ToolSpec]:
    specs = all_specs(container)
    require_known_names(specs, container)
    return enabled_specs(specs, container.settings.tool_policy)
