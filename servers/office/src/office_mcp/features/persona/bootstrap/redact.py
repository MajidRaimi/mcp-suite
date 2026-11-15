from __future__ import annotations

import re
from collections.abc import Iterable

URL_MARK = "[url]"
EMAIL_MARK = "[email]"
PHONE_MARK = "[phone]"
NAME_MARK = "[name]"
MIN_PHONE_DIGITS = 7
URL = re.compile(r"(?:https?://|www\.)\S+", re.IGNORECASE)
EMAIL = re.compile(r"[\w.%+'-]+@[\w-]+(?:\.[\w-]+)+")
PHONE = re.compile(r"(?<!\w)\+?\d[\d \t().-]{5,}\d(?!\w)")
NAME_WORD = r"[A-Z](?:[a-z]+|[a-z]*['\u2019]?[A-Z][a-z]+)"
NAME_PAIR = re.compile(rf"\b{NAME_WORD}(?:[ \t]+{NAME_WORD})+\b")
UPPER_PAIR = re.compile(r"\b[A-Z]{2,}(?:[ \t]+[A-Z]{2,})+\b")
TITLED_NAME = re.compile(rf"\b(?:Mr|Mrs|Ms|Dr)\.?[ \t]+{NAME_WORD}")
HANDLE = re.compile(r"(?<![\w@])@[\w.-]*\w")
ARABIC_WORD = "[\u0621-\u064a\u0660-\u0669\u0671-\u06d3]+"
ARABIC_VOCATIVE = re.compile(rf"(?<!\w)(\u064a\u0627[ \t]+){ARABIC_WORD}(?:[ \t]+{ARABIC_WORD})?")
NAME_TOKEN = re.compile(r"\w{2,}")
NAME_STOPWORDS = frozenset({"and", "the", "of", "al", "el", "bin", "team", "chat", "group"})
ADDRESSED_NAME = re.compile(r"(?m)(^|(?<=[.!?]\s))([A-Z][a-z]+)(?=,)")
COMMON_OPENERS = frozenset(
    {
        "Best",
        "Cheers",
        "Dear",
        "Done",
        "Great",
        "Hello",
        "Hey",
        "Hi",
        "No",
        "Ok",
        "Okay",
        "Regards",
        "Sorry",
        "Sure",
        "Thanks",
        "Yes",
    }
)


def redact(text: str, names: Iterable[str] = ()) -> str:
    without_links = URL.sub(URL_MARK, text)
    without_emails = EMAIL.sub(EMAIL_MARK, without_links)
    without_phones = PHONE.sub(phone_replacement, without_emails)
    without_handles = HANDLE.sub(NAME_MARK, without_phones)
    without_known = known_names(names).sub(NAME_MARK, without_handles)
    without_titled = TITLED_NAME.sub(NAME_MARK, without_known)
    without_upper = UPPER_PAIR.sub(NAME_MARK, without_titled)
    without_pairs = NAME_PAIR.sub(NAME_MARK, without_upper)
    without_addressed = ADDRESSED_NAME.sub(addressed_replacement, without_pairs)
    return ARABIC_VOCATIVE.sub(lambda found: found.group(1) + NAME_MARK, without_addressed)


def name_tokens(names: Iterable[str]) -> list[str]:
    tokens = {
        token
        for name in names
        for token in NAME_TOKEN.findall(EMAIL.sub(" ", name))
        if token.casefold() not in NAME_STOPWORDS and not token.isdigit()
    }
    return sorted(tokens, key=len, reverse=True)


def known_names(names: Iterable[str]) -> re.Pattern[str]:
    tokens = name_tokens(names)
    if not tokens:
        return re.compile(r"(?!)")
    alternatives = "|".join(re.escape(token) for token in tokens)
    return re.compile(rf"(?<!\w)(?:{alternatives})(?!\w)", re.IGNORECASE)


def addressed_replacement(found: re.Match[str]) -> str:
    word = found.group(2)
    return found.group(1) + (word if word in COMMON_OPENERS else NAME_MARK)


def phone_replacement(found: re.Match[str]) -> str:
    digits = sum(1 for char in found.group() if char.isdigit())
    return PHONE_MARK if digits >= MIN_PHONE_DIGITS else found.group()
