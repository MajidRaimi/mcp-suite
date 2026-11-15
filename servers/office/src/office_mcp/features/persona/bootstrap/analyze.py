from __future__ import annotations

import statistics
from collections import Counter
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from itertools import chain, zip_longest

from office_mcp.features.persona.bootstrap.collect import Corpus
from office_mcp.features.persona.bootstrap.redact import redact
from office_mcp.features.persona.models import DEFAULT_EMOJI_ALLOW, Language
from office_mcp.features.persona.rules.emoji import count_emoji
from office_mcp.features.persona.rules.language import (
    MIN_CONTEXT_LETTERS,
    language_of,
    language_sample,
    letter_count,
)
from office_mcp.features.persona.rules.signoff import (
    SIGN_OFF_PATTERN,
    TRAILING_PUNCTUATION,
    is_sign_off_like,
    non_empty_lines,
)
from office_mcp.shared.drafts import DASH_CHARACTERS
from office_mcp.shared.text import normalize

TOP_ENTRIES = 5
MAX_SNIPPETS = 10
SNIPPET_CHARS = 160
MAX_GREETING_WORDS = 6
NO_CLOSING = "(no closing)"
OTHER_GREETING = "(other)"
GREETING_WORDS = frozenset(
    normalize(word)
    for word in (
        "hi",
        "hello",
        "hey",
        "dear",
        "good",
        "morning",
        "team",
        "all",
        "salam",
        "السلام",
        "مرحبا",
        "اهلا",
        "هلا",
        "صباح",
        "مساء",
        "عزيزي",
    )
)
MORE = " ..."


@dataclass(frozen=True)
class Sample:
    text: str
    replied_to: str
    people: tuple[str, ...] = ()


@dataclass(frozen=True)
class ChannelStats:
    count: int
    english: int
    arabic: int
    with_context: int
    mirrored: int
    median_words: float
    greetings: tuple[tuple[str, int], ...]
    closings: tuple[tuple[str, int], ...]
    dashes: int
    double_hyphens: int
    spaced_hyphens: int
    emoji: int


@dataclass(frozen=True)
class Stats:
    email: ChannelStats
    teams: ChannelStats
    snippets: tuple[str, ...]


def analyze(corpus: Corpus) -> Stats:
    emails = [Sample(item.body, item.replied_to, item.people) for item in corpus.emails]
    messages = [Sample(item.text, item.replied_to, item.people) for item in corpus.messages]
    return Stats(
        email=channel_stats(emails),
        teams=channel_stats(messages),
        snippets=snippets(emails, messages),
    )


def channel_stats(samples: Sequence[Sample]) -> ChannelStats:
    languages = [language_of(language_sample(item.text)) for item in samples]
    contexts = [item for item in samples if has_context(item)]
    return ChannelStats(
        count=len(samples),
        english=languages.count(Language.EN),
        arabic=languages.count(Language.AR),
        with_context=len(contexts),
        mirrored=sum(1 for item in contexts if mirrors(item)),
        median_words=median_words(samples),
        greetings=top(greeting_key(item.text) for item in samples),
        closings=top(closing_key(item.text) for item in samples),
        dashes=sum(char in DASH_CHARACTERS for item in samples for char in item.text),
        double_hyphens=sum(item.text.count("--") for item in samples),
        spaced_hyphens=sum(item.text.count(" - ") for item in samples),
        emoji=sum(count_emoji(item.text, DEFAULT_EMOJI_ALLOW) for item in samples),
    )


def has_context(sample: Sample) -> bool:
    return letter_count(language_sample(sample.replied_to)) >= MIN_CONTEXT_LETTERS


def mirrors(sample: Sample) -> bool:
    written = language_of(language_sample(sample.text))
    return written is language_of(language_sample(sample.replied_to))


def median_words(samples: Sequence[Sample]) -> float:
    counts = [len(item.text.split()) for item in samples]
    return float(statistics.median(counts)) if counts else 0.0


def greeting_key(text: str) -> str | None:
    lines = non_empty_lines(text)
    if not lines:
        return None
    words = lines[0].split()
    if len(words) > MAX_GREETING_WORDS:
        return None
    first = normalize(words[0]).strip(TRAILING_PUNCTUATION)
    if first not in GREETING_WORDS:
        return OTHER_GREETING
    return first + MORE if len(words) > 1 else first


def closing_key(text: str) -> str:
    lines = non_empty_lines(text)
    if not lines or not is_sign_off_like(lines[-1]):
        return NO_CLOSING
    folded = normalize(lines[-1]).rstrip(TRAILING_PUNCTUATION).strip()
    found = SIGN_OFF_PATTERN.match(folded)
    phrase = folded if found is None else found.group().rstrip(TRAILING_PUNCTUATION).strip()
    return phrase if phrase == folded else phrase + MORE


def top(keys: Iterable[str | None]) -> tuple[tuple[str, int], ...]:
    counted = Counter(key for key in keys if key is not None)
    return tuple(counted.most_common(TOP_ENTRIES))


def snippets(emails: Sequence[Sample], messages: Sequence[Sample]) -> tuple[str, ...]:
    paired = chain.from_iterable(zip_longest(emails, messages))
    chosen: list[str] = []
    for item in paired:
        if item is None:
            continue
        snippet = snippet_of(item.text, item.people)
        if snippet and snippet not in chosen:
            chosen.append(snippet)
        if len(chosen) == MAX_SNIPPETS:
            break
    return tuple(chosen)


def snippet_of(text: str, people: Sequence[str] = ()) -> str:
    flattened = redact(" ".join(text.split()), people)
    if len(flattened) <= SNIPPET_CHARS:
        return flattened
    return flattened[: SNIPPET_CHARS - len(MORE)].rstrip() + MORE
