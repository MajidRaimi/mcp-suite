# Live testing

The unit tests cover every tool against fakes. This protocol drives the real Outlook and Teams on the Mac, and it is self-only: every mail goes to your own address and every Teams message goes to your own chat. Placeholders in angle brackets are filled in at run time; nothing personal is written into the repository.

## Preconditions

Every item must hold, otherwise abort.

- Only one UI driver runs. Stop any other office-mcp server or tunnel first (for example `systemctl --user stop office-mcp.service` on the machine that runs the tunnel) and confirm on the Mac that `pgrep -fl office_mcp` shows nothing.
- A dev copy of the branch on the Mac (for example `~/mcp-suite-dev`), installed with `uv sync --package office-mcp`.
- Environment for the run:

  ```sh
  OFFICE_MCP_TOOLS=all
  OFFICE_MCP_EXPERIMENTAL=all
  OFFICE_MCP_DISABLED_TOOLS=calendar_respond
  OFFICE_MCP_SEND_ONLY_TO=<own address>,teams:<self thread>
  OFFICE_MCP_PERSONA_FILE=<test copy of the persona>
  OFFICE_MCP_PERSONA_REQUIRED=true
  OFFICE_MCP_TIMEZONE=<as on the live deploy>
  OFFICE_MCP_OWNER_NAME=<as on the live deploy>
  OFFICE_MCP_DOWNLOAD_DIR=~/Library/Caches/office-mcp/downloads
  OFFICE_MCP_LIVE=1
  ```

  When the self chat is not `48:notes`, also set `OFFICE_MCP_SELF_CHAT_THREAD` or `OFFICE_MCP_LIVE_SELF_THREAD` to its thread id.
- Nobody is using the Mac (the idle check stays on).
- Run id `R` is a UTC timestamp such as `20261115T120000Z`. Every subject and Teams text starts with `[mcp-suite test] R`. Export it as `OFFICE_MCP_LIVE_RUN=R` so the pytest harness uses the same id.
- Never use a real third-party target. Negative tests use `nobody@example.invalid` or `dry_run=true`.

## Steps

Record the returned record and pass or fail for each step.

0. **Baseline, read-only.** `office_status`: record `outlook_account`, `outlook_mode`, `frontmost_app` and that `persona.state` is `loaded`. `teams_chats`: find the chat whose thread id equals `OFFICE_MCP_SELF_CHAT_THREAD` or, if absent, the "(You)" chat's thread id. Set the allowlist to `<account>,teams:<that id>` and restart. Assert that `office_status.send_only_to` is exactly those two entries. Record the Outlook folder, tab, Drafts and Deleted Items counts, open compose windows (must be none), and the Teams rail and open thread.
1. **Reads.** `outlook_folders`; `outlook_conversations view=both limit=10`; `outlook_search query="R"` (zero results, not an error); `outlook_people` with your own first name; `outlook_needs_reply`; `outlook_follow_ups`; `calendar_events days=7`; `calendar_event` on one id; `teams_search` with a harmless word; `teams_activity filter=mentions limit=5`; `teams_channels`; `teams_channel_read` on a channel whose row is not unread. Pass: records returned, view restored, frontmost app unchanged, no menus or popups left.
2. **Negative, no UI.** `office_check_text` with a body containing an em dash returns a `characters` violation. `outlook_send` to yourself with that body raises PersonaViolationError, the Drafts count is unchanged and no compose window exists. `outlook_send to=["nobody@example.invalid"]` raises RecipientNotAllowedError. `teams_send` with `dry_run=true` to any other chat id from step 0 raises RecipientNotAllowedError. `teams_send` with `dry_run=true` to a channel thread id raises ForbiddenActionError.
3. **Send.** `outlook_send` to yourself, subject `[mcp-suite test] R T1`, a multi-line body without a closing. Pass: status `sent`, `persona.fixes` contains `signoff.appended` when the account matches, frontmost app restored, Drafts at baseline. If an alert (a sensitivity label or anything else) appears: record its text. SendRefusedError with "nothing was sent" is the expected behavior, and a `sensitivity` parameter becomes a must-fix.
4. **Duplicate.** The same `outlook_send` again raises DuplicateSendError.
5. **Find and read.** `outlook_search query="R T1"` returns it with `folder`. `outlook_read` with its key returns the body exactly (closing included), `messages` has length 1, the view is not in search mode afterwards, and unread is restored if it was unread.
6. **Mark and flag.** `outlook_mark read=false`, then `read=true` (verified). `outlook_flag state=today`, then `state=clear` (verified).
7. **Reply family.** `outlook_reply` to T1; `outlook_reply reply_all=true` (pass also needs the result `recipients` to be only you); `outlook_forward` to yourself. Each `sent`.
8. **Undo.** `outlook_send` T2 to yourself, then `outlook_undo_send` at once. Pass: result ok, T2 absent from Sent Items after 30 seconds, a draft copy in Deleted Items.
9. **Move (experimental).** `outlook_move` T1 to `Archive`, then back to `Inbox`; `outlook_move destination=other`, then `focused`. Each verified.
10. **Archive and delete.** `outlook_archive` T1 (row action path, unread state kept); `outlook_delete` T1 from Archive (in Deleted Items afterwards); `outlook_delete` on it inside Deleted Items raises ForbiddenActionError.
11. **Attachment (experimental).** Only when a mail you sent with an attachment already exists in Sent Items: `outlook_save_attachment` saves into the download folder; delete the saved file afterwards. Otherwise record "not tested" and keep the tool experimental.
12. **Compose abort.** `uv run pytest -m live tests/live/test_compose_abort.py`. It opens a compose window and stops after the body is written (`fail_at="after_body"`). Pass: the window is closed, Drafts at baseline, and the error mentions the copy in Deleted Items.
13. **Teams writes, self chat only.** `teams_send` single line; `teams_send` with two lines (one message, both lines); the same send again raises DuplicateSendError; `teams_react reaction=like` (state on), the same call again (`changed=false`), then `remove=true`; `teams_reply` quoting the first message (a new own message with a quote); `teams_mark_unread` raises ActionNotVerifiedError cleanly; `teams_delete_message` on every test message of run R.
14. **Cleanup.** `outlook_search query="R" scope=all`; `outlook_delete` every hit not already in Deleted Items (Inbox, Sent Items, Archive). Items left in Deleted Items are listed in the report and stay there by design.
15. **Postflight.** `office_status` and a manual look: frontmost app, Outlook module, folder, tab, no compose windows, no menus or sheets, Drafts at baseline, Teams rail and open thread as at baseline, `teams_enhanced_mode` back to its baseline value. Restart any stopped server only after the new build is deployed or the run is abandoned.

## What counts as pass

Every step passes as written; no message reached anyone but you (check the Sent Items recipients, and that the self chat is the only Teams chat with new messages); no permanent delete happened; the UI matches the baseline at step 15.

A failing experimental step keeps that tool experimental and is reported. A failing stable step blocks the release until it is fixed and the run is repeated from step 0. Experimental tools that passed (`outlook_move`, `outlook_undo_send`, and `outlook_save_attachment` when tested) are then switched to stable. `teams_mark_unread` stays experimental until it is verified by hand on a chat with messages from others, and `calendar_respond` stays experimental and disabled.

## Harness

`tests/live/` holds the pytest part of the protocol. It is excluded from normal runs by the `live` marker (`addopts = "-m 'not live'"`), skips unless `OFFICE_MCP_LIVE=1`, and aborts the session unless `office_status.send_only_to` is exactly `[<outlook_account>, teams:<self thread>]`. Keep the run report outside the repository.
