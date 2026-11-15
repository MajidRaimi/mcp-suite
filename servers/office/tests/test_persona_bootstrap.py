from __future__ import annotations

import stat
from pathlib import Path

import pytest
from fakes.p7_persona import FakeOwnChatSource, FakeSentMailSource

from office_mcp.features.persona.bootstrap.analyze import analyze, closing_key, greeting_key
from office_mcp.features.persona.bootstrap.collect import Corpus, collect_corpus
from office_mcp.features.persona.bootstrap.redact import redact
from office_mcp.features.persona.bootstrap.report import CORPUS_NAME, STATS_NAME, write_report
from office_mcp.features.persona.bootstrap.sources import OwnMessage, SentMail
from office_mcp.shared.errors import PersonaError

ARABIC_MESSAGE = "تمام، سأرسل الملف غدا"
ARABIC_CONTEXT = "هل انتهيت من ميزانية المشروع"
MAILS = [
    SentMail(
        subject="Project Atlas budget",
        body="Hi Jane,\n\nThe budget is approved. Call me on +1 415 555 0100 or see https://contoso.com/atlas.\n\nBest regards,",
        replied_to="Could you confirm the Project Atlas budget?",
    ),
    SentMail(
        subject="Invoice", body="Omar,\n\nSent to omar@example.com today \u2014 done.\n\nThanks,"
    ),
    SentMail(
        subject="Invoice", body="Omar,\n\nSent to omar@example.com today \u2014 done.\n\nThanks,"
    ),
]
MESSAGES = [
    OwnMessage(
        chat="19:abc@unq.gbl.spaces",
        text="Done, uploaded \U0001f44d",
        replied_to="Can you upload the file?",
    ),
    OwnMessage(chat="19:abc@unq.gbl.spaces", text=ARABIC_MESSAGE, replied_to=ARABIC_CONTEXT),
    OwnMessage(chat="19:def@unq.gbl.spaces", text="Will do -- after lunch"),
]


@pytest.fixture
def home(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    home = tmp_path / "home"
    home.mkdir()
    monkeypatch.setenv("HOME", str(home))
    return home


def corpus() -> Corpus:
    return collect_corpus(FakeSentMailSource(MAILS), FakeOwnChatSource(MESSAGES), 40, 15, 20)


def test_collect_dedupes_and_respects_limits() -> None:
    mail = FakeSentMailSource(MAILS)
    chats = FakeOwnChatSource(MESSAGES)

    collected = collect_corpus(mail, chats, emails=40, chats_limit=15, per_chat=20)

    assert len(collected.emails) == 2
    assert len(collected.messages) == 3
    assert mail.limits == [40]
    assert chats.calls == [(15, 20)]


def test_collect_with_zero_skips_the_sources() -> None:
    mail = FakeSentMailSource(MAILS)
    chats = FakeOwnChatSource(MESSAGES)

    collected = collect_corpus(mail, chats, emails=0, chats_limit=0, per_chat=20)

    assert collected == Corpus(emails=(), messages=())
    assert mail.limits == []
    assert chats.calls == []


def test_collect_refuses_out_of_range_counts() -> None:
    with pytest.raises(PersonaError, match="emails must be between 0 and 200"):
        collect_corpus(FakeSentMailSource([]), FakeOwnChatSource([]), 500, 1, 1)


def test_redacts_emails_phones_urls_and_name_pairs() -> None:
    text = "Ask Jane Doe at jane@example.com or +1 555 010 4567, see https://contoso.com/x and www.example.com"

    assert redact(text) == "[name] at [email] or [phone], see [url] and [url]"


def test_redacts_a_single_addressed_name() -> None:
    assert (
        redact("Omar,\nSent it. Jane, see above. Thanks, all")
        == "[name],\nSent it. [name], see above. Thanks, all"
    )


def test_redacts_handles_titles_upper_pairs_and_mixed_case_names() -> None:
    text = "Ping @Bob. Mr. O'Brien and ALICE EXAMPLE. We met McDonald Smith."

    assert redact(text) == "Ping [name]. [name] and [name]. We met [name]."


def test_redacts_known_participants_in_both_scripts() -> None:
    arabic = "\u0645\u062d\u0645\u062f \u0627\u0644\u0639\u062a\u064a\u0628\u064a"
    text = f"ask alice later, and {arabic} agreed"

    redacted = redact(text, ("Alice Example", arabic, "bob@example.com"))

    assert redacted == "ask [name] later, and [name] [name] agreed"


def test_redacts_an_arabic_vocative_name() -> None:
    text = "\u064a\u0627 \u0645\u062d\u0645\u062f \u0627\u0644\u0639\u062a\u064a\u0628\u064a \u062a\u0645"

    assert redact(text).startswith("\u064a\u0627 [name]")
    assert "\u0645\u062d\u0645\u062f" not in redact(text)


def test_redaction_keeps_short_numbers_and_single_words() -> None:
    assert redact("Meet at 10:30 in room 204, Jane") == "Meet at 10:30 in room 204, Jane"


def test_greeting_and_closing_keys_hide_names() -> None:
    assert greeting_key("Hi Jane,\nbody") == "hi ..."
    assert greeting_key("Omar,\nbody") == "(other)"
    assert greeting_key("Hello!") == "hello"
    assert closing_key("body\n\nBest regards,") == "best regards"
    assert closing_key("body\n\nThanks, Jane") == "thanks ..."
    assert closing_key("body only") == "(no closing)"


def test_analyze_counts_languages_mirroring_dashes_and_emoji() -> None:
    stats = analyze(corpus())

    assert (stats.email.count, stats.email.english, stats.email.arabic) == (2, 2, 0)
    assert (stats.teams.count, stats.teams.english, stats.teams.arabic) == (3, 2, 1)
    assert (stats.teams.with_context, stats.teams.mirrored) == (2, 2)
    assert stats.email.dashes == 1
    assert stats.teams.double_hyphens == 1
    assert stats.teams.emoji == 1
    assert ("best regards", 1) in stats.email.closings
    assert len(stats.snippets) <= 10
    assert not any("Omar" in snippet or "Jane" in snippet for snippet in stats.snippets)


def test_writes_stats_with_mode_600_and_redacted_snippets(home: Path) -> None:
    out = home / ".config" / "office-mcp"

    path = write_report(corpus(), out, keep_corpus=False)

    assert path == (out / STATS_NAME).resolve()
    assert stat.S_IMODE(path.stat().st_mode) == 0o600
    assert stat.S_IMODE(out.stat().st_mode) == 0o700
    content = path.read_text(encoding="utf-8")
    assert "jane@example.com" not in content
    assert "omar@example.com" not in content
    assert "555 0100" not in content
    assert "contoso.com" not in content
    assert "[email]" in content
    assert "[U+2014]" in content
    assert "\u2014" not in content
    assert not (out / CORPUS_NAME).exists()
    assert sorted(item.name for item in out.iterdir()) == [STATS_NAME]


def test_keep_corpus_writes_the_corpus_with_mode_600(home: Path) -> None:
    out = home / ".config" / "office-mcp"

    write_report(corpus(), out, keep_corpus=True)

    corpus_path = out / CORPUS_NAME
    assert stat.S_IMODE(corpus_path.stat().st_mode) == 0o600
    assert len(corpus_path.read_text(encoding="utf-8").splitlines()) == 5


def test_existing_stats_file_is_made_private(home: Path) -> None:
    out = home / ".config" / "office-mcp"
    out.mkdir(parents=True)
    (out / STATS_NAME).write_text("old")
    (out / STATS_NAME).chmod(0o644)

    path = write_report(corpus(), out, keep_corpus=False)

    assert stat.S_IMODE(path.stat().st_mode) == 0o600
    assert "old" not in path.read_text(encoding="utf-8")


def test_subdirectory_is_allowed(home: Path) -> None:
    path = write_report(corpus(), home / ".config" / "office-mcp" / "runs", keep_corpus=False)

    assert path.parent.name == "runs"


@pytest.mark.parametrize(
    "relative", [".config", "Documents", ".config/office-mcp-other", ".config/office-mcp/../other"]
)
def test_refuses_an_out_dir_outside_the_config_directory(home: Path, relative: str) -> None:
    with pytest.raises(PersonaError, match="writes only inside"):
        write_report(corpus(), home / relative, keep_corpus=False)

    assert not (home / relative / STATS_NAME).exists()


def test_refuses_a_symlink_that_escapes(home: Path, tmp_path: Path) -> None:
    outside = tmp_path / "outside"
    outside.mkdir()
    (home / ".config").mkdir()
    (home / ".config" / "office-mcp").symlink_to(outside)

    with pytest.raises(PersonaError, match="writes only inside"):
        write_report(corpus(), home / ".config" / "office-mcp", keep_corpus=False)

    assert list(outside.iterdir()) == []


def test_refuses_inside_a_git_work_tree(home: Path) -> None:
    (home / ".git").mkdir()

    with pytest.raises(PersonaError, match="git work tree"):
        write_report(corpus(), home / ".config" / "office-mcp", keep_corpus=False)

    assert not (home / ".config" / "office-mcp").exists()
