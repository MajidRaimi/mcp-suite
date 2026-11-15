from __future__ import annotations

import argparse
import inspect
import io
import json
import sys
import types
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any, Literal, NoReturn, Union, get_args, get_origin, get_type_hints

from mcp.server.mcpserver.exceptions import ToolError

from office_mcp.bootstrap import BootstrapRequest, run_bootstrap
from office_mcp.container import Container, all_specs, build_container, require_known_names
from office_mcp.shared.errors import OfficeError, ToolDisabledError
from office_mcp.shared.registry import ToolSpec
from office_mcp.shared.settings import Settings

AddParser = Callable[..., argparse.ArgumentParser]
TOOL = "tool"
BOOTSTRAP = "persona-bootstrap"
SCALARS: tuple[type, ...] = (str, int, float)
ALIASES = {
    "office_status": "status",
    "outlook_folders": "folders",
    "outlook_conversations": "mail",
    "outlook_read": "read-mail",
    "teams_chats": "chats",
    "teams_read": "read-chat",
    "teams_needs_reply": "needs-reply",
}
LEGACY_FLAGS = {"outlook_read": (("--mark-read", "keep_unread"),)}
DISABLED = (
    "{name} is turned off here; enable it with OFFICE_MCP_TOOLS, OFFICE_MCP_EXPERIMENTAL or "
    "OFFICE_MCP_DISABLED_TOOLS (OFFICE_MCP_READ_ONLY=true hides every write tool)"
)


def command_name(spec: ToolSpec) -> str:
    return spec.name.replace("_", "-")


def flag(name: str) -> str:
    return "--" + name.replace("_", "-")


def without_none(hint: object) -> tuple[object, bool]:
    if get_origin(hint) not in (Union, types.UnionType):
        return hint, False
    members = [member for member in get_args(hint) if member is not type(None)]
    if len(members) != 1:
        raise TypeError(f"unsupported union {hint!r}")
    return members[0], True


def add_list(
    parser: argparse.ArgumentParser, name: str, hint: object, default: object, required: bool
) -> None:
    (item,) = get_args(hint) or (str,)
    if item not in SCALARS:
        raise TypeError(f"unsupported list item {item!r} for {name}")
    parser.add_argument(
        flag(name),
        dest=name,
        action="append",
        type=item,
        required=required,
        default=None if required else default,
        metavar=name.upper(),
    )


def add_bool(parser: argparse.ArgumentParser, name: str, default: object, required: bool) -> None:
    parser.add_argument(
        flag(name),
        dest=name,
        action=argparse.BooleanOptionalAction,
        required=required,
        default=None if required else default,
    )


def add_scalar(
    parser: argparse.ArgumentParser,
    name: str,
    kind: type,
    choices: tuple[object, ...] | None,
    default: object,
    required: bool,
) -> None:
    if required:
        parser.add_argument(name, type=kind, choices=choices)
        return
    parser.add_argument(flag(name), dest=name, type=kind, choices=choices, default=default)


def literal_kind(hint: object) -> tuple[type, tuple[object, ...]]:
    choices = get_args(hint)
    kinds = {type(choice) for choice in choices}
    if len(kinds) != 1 or next(iter(kinds)) not in SCALARS:
        raise TypeError(f"unsupported literal {hint!r}")
    return next(iter(kinds)), choices


def add_parameter(
    parser: argparse.ArgumentParser, name: str, hint: object, default: object
) -> None:
    required = default is inspect.Parameter.empty
    base, _optional = without_none(hint)
    origin = get_origin(base)
    if base is bool:
        add_bool(parser, name, default, required)
    elif origin is list:
        add_list(parser, name, base, default, required)
    elif origin is Literal:
        kind, choices = literal_kind(base)
        add_scalar(parser, name, kind, choices, default, required)
    elif isinstance(base, type) and base in SCALARS:
        add_scalar(parser, name, base, None, default, required)
    else:
        raise TypeError(f"unsupported parameter {name}: {hint!r}")


def add_legacy_flags(parser: argparse.ArgumentParser, spec: ToolSpec) -> None:
    for legacy, dest in LEGACY_FLAGS.get(spec.name, ()):
        parser.add_argument(legacy, dest=dest, action="store_false", help=argparse.SUPPRESS)


def summary(description: str) -> str:
    return description.split(". ")[0].replace("%", "%%")


def add_tool(add_parser: AddParser, spec: ToolSpec) -> None:
    alias = ALIASES.get(spec.name)
    parser = add_parser(
        command_name(spec), aliases=[alias] if alias else [], help=summary(spec.description)
    )
    hints = get_type_hints(spec.function)
    for name, parameter in inspect.signature(spec.function).parameters.items():
        add_parameter(parser, name, hints[name], parameter.default)
    add_legacy_flags(parser, spec)
    parser.set_defaults(**{TOOL: spec.name})


def add_bootstrap(add_parser: AddParser) -> None:
    parser = add_parser(BOOTSTRAP, help="Collect writing statistics from sent mail and read chats")
    parser.add_argument("--emails", type=int, default=40)
    parser.add_argument("--chats", type=int, default=15)
    parser.add_argument("--per-chat", dest="per_chat", type=int, default=20)
    parser.add_argument("--out", type=Path, default=Path("~/.config/office-mcp"))
    parser.add_argument("--keep-corpus", dest="keep_corpus", action="store_true")
    parser.set_defaults(**{TOOL: BOOTSTRAP})


def build_parser(specs: Sequence[ToolSpec]) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="office", description="Run the office MCP tools")
    commands = parser.add_subparsers(dest="command", required=True)
    for spec in specs:
        add_tool(commands.add_parser, spec)
    add_bootstrap(commands.add_parser)
    return parser


def owner_address(container: Container) -> Callable[[], str | None]:
    def address() -> str | None:
        return container.settings.owner_email or container.outlook.get().gateway.account()

    return address


def bootstrap(container: Container, arguments: dict[str, Any]) -> object:
    request = BootstrapRequest(
        emails=arguments["emails"],
        chats=arguments["chats"],
        per_chat=arguments["per_chat"],
        out=arguments["out"],
        keep_corpus=arguments["keep_corpus"],
    )
    return run_bootstrap(
        request,
        container.runner,
        container.mail.get,
        container.teams_reading.get,
        owner_address(container),
    )


def run_tool(container: Container, spec: ToolSpec, arguments: dict[str, Any]) -> object:
    if not container.settings.tool_policy.allows(spec):
        raise ToolDisabledError(DISABLED.format(name=spec.name))
    return spec.function(**arguments)


def dispatch(container: Container, specs: Sequence[ToolSpec], arguments: dict[str, Any]) -> object:
    arguments.pop("command")
    name = arguments.pop(TOOL)
    if name == BOOTSTRAP:
        return bootstrap(container, arguments)
    spec = next(spec for spec in specs if spec.name == name)
    return run_tool(container, spec, arguments)


def fail(error: Exception) -> NoReturn:
    print(json.dumps({"error": str(error)}, ensure_ascii=False))
    raise SystemExit(1) from error


def main(argv: list[str] | None = None) -> None:
    if isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout.reconfigure(errors="backslashreplace")
    try:
        container = build_container(Settings.from_env())
        specs = all_specs(container)
        require_known_names(specs, container)
    except OfficeError as error:
        fail(error)
    arguments = vars(build_parser(specs).parse_args(argv))
    try:
        result = dispatch(container, specs, arguments)
    except (OfficeError, ToolError, OSError) as error:
        fail(error)
    print(json.dumps(result, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
