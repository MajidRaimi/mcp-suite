from __future__ import annotations

from office_mcp.shared.owner import Owner


def office_persona(owner: Owner) -> str:
    return (
        f"Return the persona that describes how {owner.label} writes: the live file state, "
        "version, the prose about voice and style, the machine rules, and the same rules in "
        "words for email, Teams and RSVP comments. Read it before writing anything that will "
        f"be sent as {owner.label}, and again when a send result shows a new persona_version. "
        "Reads a local file only and changes nothing in Outlook or Teams."
    )


def office_check_text(owner: Owner) -> str:
    return (
        "Check a draft against the persona rules without sending anything. `channel` is "
        "email, teams or rsvp; `subject` is the email subject; `context` is the newest "
        "message from the other side, used to check the reply language; `account` is the "
        "sending email address. Returns ok, the final text exactly as a send tool would send "
        "it (for example with the configured closing added), every violation with a hint, the "
        f"fixes applied and warnings. Use it to fix a draft before sending as {owner.label}. "
        "Changes nothing in Outlook or Teams."
    )
