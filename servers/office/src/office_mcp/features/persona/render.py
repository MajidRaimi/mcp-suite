from __future__ import annotations

import unicodedata

from office_mcp.features.persona.models import (
    DEFAULT_SIGN_OFF_KEY,
    ChannelRules,
    CharacterSet,
    ForbidRules,
    Language,
    LanguageRules,
    PersonaState,
    Rules,
    SignOffRule,
    code_point,
)
from office_mcp.features.persona.store import PersonaSnapshot, PersonaStore
from office_mcp.shared.drafts import Channel
from office_mcp.shared.owner import Owner

MAX_INSTRUCTION_CHARS = 4000
TRUNCATION_MARK = "... (call office_persona for the rest)"
NO_PERSONA = "No persona is configured, so only the general rules apply."
NO_RULES = "no hard rules are configured"
CHECK_FIRST = "Call office_check_text first when unsure."
LANGUAGE_NAMES = {Language.EN: "English", Language.AR: "Arabic"}
CHANNEL_NAMES = {Channel.EMAIL: "Email", Channel.TEAMS: "Teams", Channel.RSVP: "RSVP comments"}
SET_WORDS = {
    CharacterSet.DASHES: (
        "no em dash, en dash, minus sign or any other dash character (use a comma, period "
        "or colon instead)"
    )
}


def instructions_block(
    store: PersonaStore, owner: Owner, max_chars: int = MAX_INSTRUCTION_CHARS
) -> str:
    snapshot = store.snapshot()
    state = snapshot.status.state
    if state is PersonaState.NONE:
        return cap(NO_PERSONA, max_chars)
    if snapshot.persona is None:
        return cap(blocked_text(state), max_chars)
    rules = snapshot.persona.rules
    rules_text = " ".join(
        f"{CHANNEL_NAMES[channel]}: {words_for(rules, channel)}."
        for channel in (Channel.EMAIL, Channel.TEAMS)
    )
    lead = (
        f"Everything you write through this server is sent as {owner.label} at once, with no "
        f"draft step. Write in {owner.possessive} voice as described below. These hard rules "
        "are checked before anything is sent, and a send that breaks one is rejected with the "
        f"list of problems: {rules_text} Call office_persona for the current persona and "
        "office_check_text to test a draft before sending."
    )
    prose = snapshot.persona.prose.strip()
    if not prose:
        return cap(lead, max_chars)
    return with_prose(f"{lead} Persona: ", prose, max_chars)


def with_prose(lead: str, prose: str, max_chars: int) -> str:
    if len(lead) + len(prose) <= max_chars:
        return lead + prose
    room = max_chars - len(lead) - len(TRUNCATION_MARK)
    if room <= 0:
        return cap(lead, max_chars)
    return lead + prose[:room].rstrip() + TRUNCATION_MARK


def cap(text: str, max_chars: int) -> str:
    return text if len(text) <= max_chars else text[:max_chars]


def blocked_text(state: PersonaState) -> str:
    return f"The persona file is {state.value}; sends are blocked until it is fixed."


def rules_in_words(store: PersonaStore, channel: Channel) -> str:
    return snapshot_words(store.snapshot(), channel)


def send_rules_text(store: PersonaStore, channel: Channel) -> str:
    return f"Persona rules for this text: {rules_in_words(store, channel)}. {CHECK_FIRST}"


def snapshot_words(snapshot: PersonaSnapshot, channel: Channel) -> str:
    if snapshot.persona is None:
        return f"the persona file is {snapshot.status.state.value}, so sends are blocked"
    return words_for(snapshot.persona.rules, channel)


def words_for(rules: Rules, channel: Channel) -> str:
    channel_rules = rules.channel(channel)
    parts = [
        *forbid_words(rules.forbid, channel),
        *([] if channel_rules is None else sign_off_words(channel_rules, channel)),
        *([] if rules.language is None else [language_words(rules.language)]),
        *([] if channel_rules is None else length_words(channel_rules)),
    ]
    return "; ".join(parts) if parts else NO_RULES


def forbid_words(forbid: ForbidRules, channel: Channel) -> list[str]:
    parts = [SET_WORDS[name] for name in forbid.character_sets]
    if forbid.characters:
        listed = ", ".join(describe(char) for char in forbid.characters)
        parts.append(f"none of these characters: {listed}")
    if forbid.emoji:
        emoticons = " and no emoticons such as :) or <3" if channel is Channel.TEAMS else ""
        parts.append(f"no emojis{emoticons}")
    if forbid.phrases:
        parts.append("never the phrases " + ", ".join(f'"{item}"' for item in forbid.phrases))
    if forbid.patterns:
        listed = ", ".join(f'"{item.pattern}"' for item in forbid.patterns)
        parts.append(f"nothing matching the regular expressions {listed}")
    return parts


def sign_off_words(rules: ChannelRules, channel: Channel) -> list[str]:
    if rules.sign_off_rule is SignOffRule.FORBIDDEN:
        return ["no closing line such as Thanks or Best regards at the end"]
    closing = closing_words(rules)
    if not closing:
        return []
    scope = accounts_words(rules) if channel is Channel.EMAIL else ""
    parts: list[str] = []
    if rules.sign_off_rule is SignOffRule.REQUIRED:
        parts.append(
            f"end with {closing} on its own line{scope} (it is added when the text has no "
            "closing, and a different closing is rejected)"
        )
    if rules.nothing_after_sign_off:
        parts.append(
            f"nothing after {closing}, not even a name, because the signature card is added "
            "automatically"
        )
    return parts


def closing_words(rules: ChannelRules) -> str:
    default = rules.sign_off.get(DEFAULT_SIGN_OFF_KEY)
    named = [
        f'"{text}" for {LANGUAGE_NAMES[Language(key)]} text'
        for key, text in rules.sign_off.items()
        if key != DEFAULT_SIGN_OFF_KEY
    ]
    if default is None:
        return " or ".join(named)
    if not named:
        return f'"{default}"'
    return f'"{default}" (' + ", ".join(named) + ")"


def accounts_words(rules: ChannelRules) -> str:
    if not rules.applies_to_accounts:
        return ""
    return " when sending from " + " or ".join(rules.applies_to_accounts)


def language_words(rules: LanguageRules) -> str:
    default = LANGUAGE_NAMES[rules.default]
    mirrored = sorted(LANGUAGE_NAMES[item] for item in rules.mirror if item is not rules.default)
    text = f"write in {default}"
    if mirrored:
        text += ", and in " + " or ".join(mirrored) + " when the other side writes in it"
    if not rules.enforce:
        text += " (a preference, not enforced)"
    return text


def length_words(rules: ChannelRules) -> list[str]:
    limits = [
        *([] if rules.max_words is None else [f"{rules.max_words} words"]),
        *([] if rules.max_chars is None else [f"{rules.max_chars} characters"]),
    ]
    if not limits:
        return []
    return ["aim for at most " + " and ".join(limits) + " (a warning, not a rejection)"]


def describe(char: str) -> str:
    name = unicodedata.name(char, "")
    return f"{code_point(char)} {name}".strip()
