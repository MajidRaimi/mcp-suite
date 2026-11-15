from __future__ import annotations

import ast
import io
import sys
import tokenize
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from pathlib import Path

PYTHON_SUFFIXES = frozenset({".py"})
HASH_SUFFIXES = frozenset({".toml", ".yaml", ".yml", ".sh", ".cfg", ".ini", ".env", ".service"})
MARKUP_SUFFIXES = frozenset({".md", ".html", ".xml", ".plist"})
HASH_NAMES = frozenset({".gitignore", ".env.example", "Dockerfile", "Makefile"})
SKIPPED_DIRECTORIES = frozenset(
    {".venv", "__pycache__", ".mypy_cache", ".ruff_cache", ".pytest_cache", ".git"}
)
DOCSTRING_OWNERS = (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)
MARKUP_OPEN = "<!--"
QUOTES = frozenset({'"', "'"})


@dataclass(frozen=True)
class Finding:
    path: Path
    line: int
    kind: str

    def render(self) -> str:
        return f"{self.path}:{self.line}: {self.kind}"


def python_findings(path: Path, source: str) -> Iterator[Finding]:
    tokens = tokenize.generate_tokens(io.StringIO(source).readline)
    for token in tokens:
        if token.type == tokenize.COMMENT and not is_shebang(token.start[0], token.string):
            yield Finding(path, token.start[0], "comment")
    tree = ast.parse(source, filename=str(path))
    for node in ast.walk(tree):
        if isinstance(node, ast.Expr) and is_string(node.value):
            kind = "docstring" if is_docstring(tree, node) else "string statement"
            yield Finding(path, node.lineno, kind)


def is_string(node: ast.expr) -> bool:
    return isinstance(node, ast.Constant) and isinstance(node.value, str)


def is_docstring(tree: ast.Module, statement: ast.Expr) -> bool:
    return any(
        isinstance(owner, DOCSTRING_OWNERS) and owner.body and owner.body[0] is statement
        for owner in ast.walk(tree)
    )


def is_shebang(line: int, text: str) -> bool:
    return line == 1 and text.startswith("#!")


def hash_findings(path: Path, source: str) -> Iterator[Finding]:
    for number, line in enumerate(source.splitlines(), start=1):
        if is_shebang(number, line):
            continue
        if hash_comment_start(line) is not None:
            yield Finding(path, number, "comment")


def hash_comment_start(line: str) -> int | None:
    quote: str | None = None
    for index, char in enumerate(line):
        if quote is not None:
            if char == quote:
                quote = None
        elif char in QUOTES:
            quote = char
        elif char == "#" and (index == 0 or line[index - 1].isspace()):
            return index
    return None


def markup_findings(path: Path, source: str) -> Iterator[Finding]:
    for number, line in enumerate(source.splitlines(), start=1):
        if MARKUP_OPEN in line:
            yield Finding(path, number, "comment")


def findings_for(path: Path) -> Iterator[Finding]:
    source = path.read_text(encoding="utf-8")
    if path.suffix in PYTHON_SUFFIXES:
        yield from python_findings(path, source)
    elif path.suffix in HASH_SUFFIXES or path.name in HASH_NAMES:
        yield from hash_findings(path, source)
    elif path.suffix in MARKUP_SUFFIXES:
        yield from markup_findings(path, source)


def is_checked(path: Path) -> bool:
    return (
        path.suffix in PYTHON_SUFFIXES | HASH_SUFFIXES | MARKUP_SUFFIXES or path.name in HASH_NAMES
    )


def files_under(paths: Iterable[Path]) -> Iterator[Path]:
    for path in paths:
        if path.is_dir():
            for child in sorted(path.rglob("*")):
                skipped = any(part in SKIPPED_DIRECTORIES for part in child.parts)
                if child.is_file() and not skipped and is_checked(child):
                    yield child
        elif path.is_file():
            yield path


def main(arguments: list[str]) -> int:
    if not arguments:
        print("usage: no_comments.py PATH [PATH ...]", file=sys.stderr)
        return 2
    paths = [Path(arg) for arg in arguments]
    missing = [path for path in paths if not path.exists()]
    for path in missing:
        print(f"{path}: no such file or directory", file=sys.stderr)
    if missing:
        return 2
    findings = [item for path in files_under(paths) for item in findings_for(path)]
    for finding in findings:
        print(finding.render())
    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
