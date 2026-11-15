from __future__ import annotations

from office_mcp.shared.owner import Owner
from office_mcp.shared.phrases import FOREGROUND_NOTE, INJECTION_RULE

SCREEN_NOTE = (
    "Teams switches to its Calendar for a moment (Week view) and then goes back to the app, "
    "view, week and chat it showed before."
)


def calendar_events(owner: Owner) -> str:
    return (
        f"List events from {owner.possessive} Teams calendar for `days` days (1 to 31) from "
        "`start` (ISO date, default today). Each event has an `id` for calendar_event and "
        "calendar_respond, subject, ISO start and end with the Mac's zone offset (dates only "
        "for all-day events), location, whether it is an online meeting, organizer, show-as "
        "(free, tentative, busy, out_of_office, working_elsewhere), recurring and canceled. "
        "Canceled events are left out unless include_canceled is true. Ids stay valid until the "
        f"server restarts. {SCREEN_NOTE} {INJECTION_RULE}"
    )


def calendar_event(owner: Owner) -> str:
    return (
        f"Show one event from {owner.possessive} Teams calendar by the `id` from "
        "calendar_events: time, recurrence, location, organizer, the invitation line, "
        f"{owner.possessive} response (accepted, tentative, declined, not_responded, organizer "
        "or unknown), response counts, notes and whether a Join button exists. Opens the "
        "event card in Teams and closes it again; never joins the meeting or opens its chat. "
        f"{SCREEN_NOTE} {INJECTION_RULE}"
    )


CHECK_RSVP = "Call office_check_text with channel rsvp first when unsure."


def calendar_respond(owner: Owner, rules: str = "") -> str:
    return (
        f"Accept, tentatively accept or decline an invitation in {owner.possessive} Teams "
        "calendar by the `id` from calendar_events. With notify=true the organizer is emailed, "
        "with the optional `comment` (checked against the persona as channel rsvp, at most "
        "2000 characters); notify=false answers without emailing and takes no comment. "
        "Answering only this occurrence: when Teams asks about a whole series nothing is sent. "
        "Giving the response the event already has changes nothing. When "
        "OFFICE_MCP_SEND_ONLY_TO is set only notify=false is allowed. dry_run=true runs every "
        "check and sends nothing. A result with verified=false may still have been sent: check "
        "with calendar_event before retrying. Use channel rsvp with office_check_text. "
        f"{FOREGROUND_NOTE} {SCREEN_NOTE} {INJECTION_RULE} {rules or CHECK_RSVP}"
    )
