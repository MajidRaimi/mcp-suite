from __future__ import annotations

import json
import os
from collections.abc import Iterator, Sequence
from pathlib import Path

from office_mcp.features.persona.bootstrap.analyze import ChannelStats, Stats, analyze
from office_mcp.features.persona.bootstrap.collect import Corpus
from office_mcp.shared.drafts import show_dashes
from office_mcp.shared.errors import PersonaError

ALLOWED_PARTS = (".config", "office-mcp")
STATS_NAME = "persona-stats.md"
CORPUS_NAME = "persona-corpus.jsonl"
DIRECTORY_MODE = 0o700
FILE_MODE = 0o600
LANGUAGE_HEADER = (
    "| Channel | Messages | English | Arabic | Replies with context | Same language as context |"
)
WRITE_FLAGS = os.O_WRONLY | os.O_CREAT | os.O_TRUNC | getattr(os, "O_NOFOLLOW", 0)


def write_report(corpus: Corpus, out_dir: Path, keep_corpus: bool) -> Path:
    target = prepare_directory(out_dir)
    stats_path = target / STATS_NAME
    write_private(stats_path, render_stats(analyze(corpus)))
    if keep_corpus:
        write_private(target / CORPUS_NAME, render_corpus(corpus))
    return stats_path


def allowed_base() -> Path:
    return Path.home().resolve().joinpath(*ALLOWED_PARTS)


def prepare_directory(out_dir: Path) -> Path:
    target = require_allowed(out_dir)
    target.mkdir(mode=DIRECTORY_MODE, parents=True, exist_ok=True)
    return require_allowed(target)


def require_allowed(out_dir: Path) -> Path:
    base = allowed_base()
    target = out_dir.expanduser().resolve()
    if target != base and base not in target.parents:
        raise PersonaError(f"the bootstrap writes only inside {base}, not {target}")
    repository = git_work_tree(target)
    if repository is not None:
        raise PersonaError(
            f"{target} is inside the git work tree {repository}; choose a directory outside "
            "any repository"
        )
    return target


def git_work_tree(path: Path) -> Path | None:
    for candidate in (path, *path.parents):
        if (candidate / ".git").exists():
            return candidate
    return None


def write_private(path: Path, content: str) -> None:
    descriptor = os.open(path, WRITE_FLAGS, FILE_MODE)
    try:
        os.fchmod(descriptor, FILE_MODE)
        os.write(descriptor, content.encode("utf-8"))
    finally:
        os.close(descriptor)


def render_corpus(corpus: Corpus) -> str:
    emails = (
        {"channel": "email", "subject": item.subject, "text": item.body} for item in corpus.emails
    )
    messages = (
        {"channel": "teams", "chat": item.chat, "text": item.text} for item in corpus.messages
    )
    records = [*emails, *messages]
    return "".join(json.dumps(record, ensure_ascii=False) + "\n" for record in records)


def render_stats(stats: Stats) -> str:
    lines = [
        "# Persona bootstrap stats",
        "",
        f"Sent emails: {stats.email.count}. Own Teams messages: {stats.teams.count}. "
        "Aggregates only; snippets are redacted.",
        "",
        "## Language",
        "",
        LANGUAGE_HEADER,
        "|---|---|---|---|---|---|",
        language_row("Email", stats.email),
        language_row("Teams", stats.teams),
        "",
        "## Length",
        "",
        f"- Email median words: {stats.email.median_words:g}",
        f"- Teams median words: {stats.teams.median_words:g}",
        "",
        *ranked("Email greetings (first word)", stats.email.greetings),
        *ranked("Email closings", stats.email.closings),
        *ranked("Teams greetings (first word)", stats.teams.greetings),
        *ranked("Teams closings", stats.teams.closings),
        "## Dashes and emoji",
        "",
        usage_line("Email", stats.email),
        usage_line("Teams", stats.teams),
        "",
        "## Snippets (redacted)",
        "",
        *(f"- {show_dashes(snippet)}" for snippet in stats.snippets),
    ]
    return "\n".join(lines).rstrip() + "\n"


def language_row(name: str, stats: ChannelStats) -> str:
    return (
        f"| {name} | {stats.count} | {stats.english} | {stats.arabic} | "
        f"{stats.with_context} | {stats.mirrored} |"
    )


def ranked(title: str, entries: Sequence[tuple[str, int]]) -> Iterator[str]:
    yield f"## {title}"
    yield ""
    if not entries:
        yield "- none"
    for key, count in entries:
        yield f"- {show_dashes(key)}: {count}"
    yield ""


def usage_line(name: str, stats: ChannelStats) -> str:
    return (
        f"- {name}: dash characters {stats.dashes}, double hyphens {stats.double_hyphens}, "
        f"spaced hyphens {stats.spaced_hyphens}, emoji {stats.emoji}"
    )
