from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path

TEMPLATE_DIR = Path(__file__).resolve().parent / "server"
SERVERS_DIR = Path(__file__).resolve().parent.parent / "servers"
NAME_PATTERN = re.compile(r"^[a-z][a-z0-9]{1,30}$")
TEMPLATE_PORT = "8770"
DEFAULT_PORT_PATTERN = re.compile(r"^DEFAULT_PORT\s*=\s*(\d+)\s*$", re.MULTILINE)
IGNORED = shutil.ignore_patterns(
    "__pycache__", "*.pyc", ".mypy_cache", ".pytest_cache", ".ruff_cache", ".venv"
)


class ScaffoldError(Exception):
    pass


def substitute(text: str, name: str, port: str) -> str:
    return (
        text.replace("__NAME__", name.upper())
        .replace("__name__", name)
        .replace(TEMPLATE_PORT, port)
    )


def claimed_ports() -> dict[str, str]:
    ports: dict[str, str] = {}
    for settings in SERVERS_DIR.glob("*/src/*/shared/settings.py"):
        for match in DEFAULT_PORT_PATTERN.finditer(settings.read_text(encoding="utf-8")):
            ports[match.group(1)] = settings.relative_to(SERVERS_DIR).parts[0]
    return ports


def validate(name: str, port: str) -> None:
    if not NAME_PATTERN.fullmatch(name):
        raise ScaffoldError(f"name must be 2 to 31 lowercase letters or digits, got {name!r}")
    if not port.isdigit() or not 1024 <= int(port) <= 65535:
        raise ScaffoldError(f"port must be between 1024 and 65535, got {port!r}")
    if (SERVERS_DIR / name).exists():
        raise ScaffoldError(f"servers/{name} already exists")
    owner = claimed_ports().get(str(int(port)))
    if owner is not None:
        raise ScaffoldError(f"port {port} is already the default port of servers/{owner}")


def rename_paths(root: Path, name: str) -> None:
    for path in sorted(root.rglob("*__name__*"), key=lambda item: len(item.parts), reverse=True):
        path.rename(path.with_name(path.name.replace("__name__", name)))


def rewrite_files(root: Path, name: str, port: str) -> None:
    for path in root.rglob("*"):
        if path.is_file():
            path.write_text(substitute(path.read_text(encoding="utf-8"), name, port), "utf-8")


def scaffold(name: str, port: str) -> Path:
    validate(name, port)
    target = SERVERS_DIR / name
    shutil.copytree(TEMPLATE_DIR, target, ignore=IGNORED)
    rename_paths(target, name)
    rewrite_files(target, name, port)
    return target


def main() -> None:
    parser = argparse.ArgumentParser(description="Create servers/<name> from templates/server")
    parser.add_argument("name")
    parser.add_argument("--port", required=True)
    args = parser.parse_args()
    try:
        target = scaffold(args.name, args.port)
    except ScaffoldError as error:
        sys.exit(f"scaffold: {error}")
    print(f"created {target.relative_to(SERVERS_DIR.parent)}")


if __name__ == "__main__":
    main()
