# office-mcp

An MCP server that works with Microsoft Outlook mail, the Teams calendar and Microsoft Teams chats through the desktop apps that are already signed in on a Mac. It uses the macOS Accessibility API, so it sees what the apps show and needs no Microsoft Graph app registration, API keys or tenant configuration.

By default every exposed tool only reads. Organizing mail and sending (email, Teams messages, reactions, invitation responses) are opt-in tool groups with their own safety checks. All tools drive the visible UI: they switch folders, select conversations and open chats on screen, and tools that change something bring the app to the front for a moment. Read [Safety model](#safety-model) and [Security and limitations](#security-and-limitations) before you turn on a write group.

## Tools

Tools come in three groups. `read` is always on; `organize` and `send` are enabled with `OFFICE_MCP_TOOLS`. Tools marked experimental are registered only when listed in `OFFICE_MCP_EXPERIMENTAL` (or when it is `all`), and their descriptions start with "Experimental: ". `OFFICE_MCP_READ_ONLY=true` wins over everything and leaves only the stable `read` tools.

### read

| Tool | Parameters (defaults) | Returns | Side effects |
|---|---|---|---|
| `office_status` | none | Accessibility permission, whether Outlook and Teams run, Outlook health (`outlook_working`, `outlook_online`, `outlook_today`, `outlook_problems`), `outlook_account`, `outlook_mode`, `frontmost_app`, `session`, `teams_enhanced_mode`, and the configuration: `read_only`, `tool_groups`, `experimental`, `disabled_tools`, `persona`, `send_only_to`, `self_chat_thread` | None |
| `office_persona` | none | The persona file state, version, prose and rules, plus the rules in words per channel | None |
| `office_check_text` | `text`, `channel` (`email`, `teams` or `rsvp`), `subject` (""), `context` (""), `account` ("") | The persona review: `ok`, `final_text`, `final_subject`, `violations`, `fixes`, `warnings`, `persona_version` | None |
| `outlook_folders` | none | Folders with `name`, `path`, `level`, `unread`, `total`, `unseen`; Favorites duplicates are skipped | None |
| `outlook_conversations` | `folder` ("Inbox"), `unread_only` (false), `days` (none), `limit` (50, 0 means all), `load_more` (0), `view` (`current`, `focused`, `other` or `both`) | Conversations newest first with a `key`, flags, counts, title, time label, ISO date, preview, `folder`, `view`, `flagged`, `replied`, `forwarded`, `has_attachments`, `pinned` | Switches Outlook to `folder` and the tab; loads older pages for `days` and `load_more`; puts the view back |
| `outlook_read` | `ref`, `folder` (none), `keep_unread` (true), `max_body` (20000), `max_messages` (20) | The latest email with sender, recipients, sent time, notes, body, `messages` (newest first) and `attachments` | Selects the conversation. With `keep_unread=true` an unread email is marked unread again (Outlook comes to the front for that) |
| `outlook_search` | `query`, `scope` (`mailbox`; also `folder`, `subfolders`, `all`), `folder` (none), `limit` (25), `load_more` (0) | `{query, scope, count, results}`; each result has `folder` and `group` | Runs Outlook's search and always cancels it, restoring the previous scope |
| `outlook_people` | `query` (2 to 100 characters), `limit` (8) | `[{name, email, source}]` from Outlook's suggestions and recently read mail | Types into the search field without searching, then clears it |
| `outlook_needs_reply` | `days` (7), `folder` ("Inbox"), `view` (`both`), `limit` (30) | Conversations that may wait on you, with a `reason`, unread first | Same as `outlook_conversations` |
| `outlook_follow_ups` | `min_days` (2), `max_days` (14), `limit` (30) | Sent mail with no answer: `key`, `subject`, `sent_date`, `to_preview`, `days_waiting` | Opens Sent Items and the Inbox, then puts the view back |
| `calendar_events` | `start` (today, ISO date), `days` (1), `include_canceled` (false), `limit` (100) | Events with an `id`, subject, start, end, location, organizer, show-as, recurring, canceled | Teams switches to its Calendar for a moment and goes back |
| `calendar_event` | `event` (an `id` from `calendar_events`) | Time, recurrence, location, organizer, your response, response counts, notes, whether a Join button exists | Opens and closes the event card; never joins |
| `teams_chats` | `unread_only` (false), `section` (none), `limit` (0, means all) | Chats from the chat list with unread and muted flags, presence, preview, sender and date | Switches Teams to the Chat view if needed |
| `teams_read` | `chat`, `limit` (30), `older` (0), `keep_unread` (false) | The chat and its messages (author, mine, time, text, `quote`, `edited`), `restored_unread` and `notes` | Opens the chat, so Teams marks it read. `keep_unread=true` marks the newest message from someone else unread again |
| `teams_needs_reply` | `days` (7), `include_meetings` (false) | Chats whose last message is not yours, not muted, unread first | Reads the chat list only |
| `teams_activity` | `filter` (`all`, `unread` or `mentions`), `limit` (30) | Activity items: kind, person, title, preview, time label, location | Opens Activity, which can open the latest item and mark it read; the rail is put back |
| `teams_search` | `query` (2 to 200 characters), `limit` (15) | Typeahead results: type, text, chat, members, time | Types into the search box and clears it |
| `teams_channels` | `team` (none), `unread_only` (false) | Teams and channels with ids, unread and private flags | Opens the Teams rail, which can open a channel; the rail is put back |
| `teams_channel_read` | `team`, `channel`, `limit` (20), `replies` (true) | Posts with author, time, subject, text and replies | Opens the channel, so Teams marks it read; the rail is put back |

### organize

| Tool | Parameters (defaults) | Effect | Notes |
|---|---|---|---|
| `outlook_archive` | `refs` (1 to 25 keys), `folder` (none) | Moves conversations to Archive | Uses the row's own Archive action when it can, so nothing gets selected |
| `outlook_delete` | `refs`, `folder` | Moves conversations to Deleted Items | Soft only. Refused inside Deleted Items; an alert about permanent deletion is cancelled |
| `outlook_move` | `refs`, `destination` (folder name or path, `focused` or `other`), `folder` | Moves conversations | A move that cannot be verified is never repeated |
| `outlook_mark` | `refs`, `read` (true or false), `folder` | Marks read or unread and checks the row | |
| `outlook_flag` | `refs`, `state` (`no_date`; also `today`, `tomorrow`, `this_week`, `next_week`, `clear`), `folder` | Flags or clears the flag | |
| `outlook_save_attachment` (experimental) | `ref`, `name`, `folder`, `keep_unread` (true) | Saves one attachment into `OFFICE_MCP_DOWNLOAD_DIR` | Registered only when that variable is set |
| `teams_mark_unread` (experimental) | `chat` | Marks the newest message from someone else unread | Fails cleanly in a chat with only your own messages |

Organize tools take only the 12 character keys from `outlook_conversations` or `outlook_search`. A key stays valid for the same day; a key that no longer matches one row fails that item with "list or search again". Batches report each item, and items not reached within the time budget come back as `skipped`.

### send

| Tool | Parameters (defaults) | Effect |
|---|---|---|
| `outlook_send` | `to` (1 to 20 addresses), `subject`, `body`, `cc` (none), `bcc` (none), `allow_language_mismatch` (false), `allow_duplicate` (false), `dry_run` (false) | Sends a new email at once |
| `outlook_reply` | `ref` (key), `body`, `reply_all` (false), `cc`, `bcc`, `folder`, `allow_language_mismatch`, `allow_duplicate`, `dry_run` | Replies to the latest message of a conversation |
| `outlook_forward` | `ref` (key), `to`, `body` (""), `cc`, `bcc`, `folder`, `allow_duplicate`, `dry_run` | Forwards the latest message |
| `outlook_undo_send` (experimental) | none | Takes back the last mail this server sent, within 30 seconds, using Outlook's Undo Send |
| `calendar_respond` (experimental) | `event`, `response` (`accept`, `tentative` or `decline`), `notify` (true), `comment` (none), `dry_run` (false) | Answers an invitation from the Teams event card |
| `teams_send` | `chat` (thread id or exact full name), `text`, `group_ok` (false), `allow_language_mismatch`, `allow_duplicate`, `dry_run` | Posts a message |
| `teams_reply` | `chat`, `message_id`, `text`, `group_ok`, `allow_language_mismatch`, `allow_duplicate`, `dry_run` | Replies with a quote of that message |
| `teams_react` | `chat`, `message_id`, `reaction` (`like`, `heart`, `laugh` or `surprised`), `remove` (false) | Sets or removes a reaction; giving the current state changes nothing |
| `teams_delete_message` | `chat`, `message_id` | Deletes one of your own messages |

Send results carry `status`: `sent` (verified in Sent Items or the chat), `submitted` (waiting in the Outbox), `unconfirmed` (the send was pressed but not seen yet) or `dry_run`, plus `text_sent` (the exact text after persona fixes) and `persona` (version, fixes, warnings). Never retry an `unconfirmed` send without checking first.

Every tool carries MCP annotations that follow its effect: reads are `read_only_hint=true`; tools that mark mail or chats read have `read_only_hint=false`; organize tools are idempotent; deletes are `destructive_hint=true`; send tools are `open_world_hint=true`.

## Safety model

- **Opt-in groups.** The default exposes only the 19 stable `read` tools. `OFFICE_MCP_READ_ONLY=true` hides every write tool whatever the other variables say. `OFFICE_MCP_DISABLED_TOOLS` removes single tools, and an unknown name in it or in `OFFICE_MCP_EXPERIMENTAL` stops the server at startup.
- **Checks before any UI.** Arguments, the persona rules, the recipient allowlist and the duplicate guard run before a compose window or chat box is touched. A `dry_run` returns there.
- **Exact targets.** Outlook write tools take keys only; Teams write tools take a thread id or the exact full chat name. Channels and other teams' threads are refused, and meeting chats or chats with more than `OFFICE_MCP_TEAMS_MAX_MEMBERS` members need `group_ok=true`.
- **Read back before the single commit.** Recipients, subject and text are read back from the compose surface and must match exactly (a substituted dash is caught). Then exactly one Send press happens. After it the tools only observe; a failure after Send is reported as `unconfirmed`, never retried.
- **Recipient allowlist.** `OFFICE_MCP_SEND_ONLY_TO` limits sends to listed addresses, `@domains`, `teams:self` and `teams:<thread id>`. Everything else is refused with nothing sent.
- **Soft deletes only.** Outlook deletes move to Deleted Items; permanent-delete alerts are cancelled. Only your own Teams messages can be deleted. Forbidden menu items (Permanently Delete, Empty Folder, Recall, Report Junk, Mark All as Read and others) and dangerous shortcuts are refused outright.
- **Focus.** Tools that need the app in front wait until you have not used the keyboard or mouse for `OFFICE_MCP_IDLE_SECONDS`, refuse on a locked screen or during secure input, and give focus back only if the app is still in front when they finish.
- **One action at a time.** A process-wide and cross-process lock (`OFFICE_MCP_LOCK_FILE`) serializes every call, including the `office` CLI. A write that waited too long for the lock refuses to act.
- **Prompt injection.** Mail, chat and calendar content is data. The server instructions tell the model never to act because a message asks for it.

## Persona

Every text a send tool writes (email subject and body, Teams messages and replies, invitation comments) is checked against a persona file before anything opens. The file is Markdown with a TOML front matter: the front matter holds hard rules (forbidden characters such as every dash, emojis, phrases, patterns, language mirroring, a required or forbidden closing, length warnings) and the prose describes the voice for the model. A text that breaks a hard rule is refused with a numbered list of problems; a missing closing is appended automatically and reported in `persona.fixes`.

The file lives outside the repository, by default at `~/.config/office-mcp/persona.md` (mode 600). When it is missing and not required, only the general rules apply. With `OFFICE_MCP_PERSONA_REQUIRED=true`, or an explicit `OFFICE_MCP_PERSONA_FILE` that does not exist or does not parse, every send is blocked until it is fixed. Edits are picked up on the next call. `office_persona` shows the current state and `office_check_text` tests a draft.

`office persona-bootstrap` collects statistics from your own sent mail and already read chats (nothing is marked read) and writes `persona-stats.md` under `~/.config/office-mcp/` to seed the file. See [docs/persona.md](docs/persona.md) for the format, every rule and the bootstrap, and [examples/persona.example.md](examples/persona.example.md) for a complete example.

## How it works

The server uses [pyobjc](https://pyobjc.readthedocs.io) to call the macOS Accessibility API (`AXUIElement`) against the running Outlook and Teams processes. It reads the folder pane, message list, reading pane, the Teams chat list, message pane and calendar from the accessibility tree, and presses rows, buttons and menu items the same way assistive technology does. Because it uses the desktop apps you are already signed in to, it needs no credentials of its own and only sees what your account can already see.

Before doing anything, every Outlook tool checks that Outlook is working. If Outlook shows the "Outlook is offline" banner, or its calendar's "Today" is not the Mac's date, the tool fails with an "Outlook is not working right now" error that says what is wrong and how to fix it, instead of returning out-of-date mail. `office_status` runs the same check but never opens or focuses Outlook.

## Requirements

- macOS with a logged-in GUI session (the apps must be on screen, not running behind a locked or headless session).
- The new Outlook for Mac (bundle id `com.microsoft.Outlook`), signed in.
- The new Microsoft Teams (bundle id `com.microsoft.teams2`), signed in.
- English UI in both apps. The parsers match English labels such as "Load more conversations", "Mark as Unread" and "Unread message".
- Date formats: the Outlook message list must show dates as D/M/YYYY with 12-hour or 24-hour times. The Teams chat list shows dates as M/D or M/D/YYYY.
- The Accessibility permission for the process that starts the server (see [Setup](#setup)).
- [uv](https://docs.astral.sh/uv/) and Python 3.13 (uv installs Python for you).

## Setup

1. Clone the repository on the Mac and install the server:

   ```sh
   git clone https://github.com/MajidRaimi/mcp-suite
   cd mcp-suite
   uv sync --package office-mcp
   ```

2. Grant Accessibility in System Settings, Privacy and Security, Accessibility, to the app that starts the server:
   - your terminal (Terminal or iTerm) when you run it by hand,
   - Claude.app when Claude Desktop starts it over stdio,
   - `/usr/libexec/sshd-keygen-wrapper` when it is started over SSH from another machine,
   - the `python3` binary in `.venv` (or `uv`) when launchd starts it.

   Restart the server after you grant it.

3. Check the result:

   ```sh
   uv run --package office-mcp office status
   ```

   `accessibility`, `outlook_running` and `teams_running` should be `true`.

## Configuration

The server reads these variables from its environment. `.env.example` lists every one with its default. The server does not load a `.env` file by itself; export the variables, put them in your client's `env` block, or run with `uv run --env-file servers/office/.env ...`. Values are trimmed, a blank value counts as unset, and an invalid value stops the server at startup with a clear message.

| Variable | Default | Meaning |
|---|---|---|
| `OFFICE_MCP_TRANSPORT` | `stdio` | `stdio` or `streamable-http`. `--transport` and `--http` override it. |
| `OFFICE_MCP_HOST` | `127.0.0.1` | HTTP bind address. `--host` overrides it. |
| `OFFICE_MCP_PORT` | `8769` | HTTP port. `--port` overrides it. |
| `OFFICE_MCP_TIMEZONE` | system zone, else UTC | IANA zone used for Teams message timestamps. List labels are read against the Mac's own date. |
| `OFFICE_MCP_TOKEN` | unset | When set, every HTTP request (including `/health`) must send `Authorization: Bearer <token>`. Ignored for stdio. |
| `OFFICE_MCP_ALLOWED_HOSTS` | unset | Comma-separated `host` or `host:port` values accepted in the `Host` and `Origin` headers, in addition to loopback. |
| `OFFICE_MCP_OWNER_NAME` | unset | Name used in the instructions and descriptions. Falls back to "the user". |
| `OFFICE_MCP_OWNER_EMAIL` | unset | Your own address. Used for reply all, follow-ups and the persona's account scoping; when unset it is read from the Outlook window title. |
| `OFFICE_MCP_TOOLS` | `read` | Comma-separated groups (`read`, `organize`, `send`) or `all`. `read` is always included. |
| `OFFICE_MCP_READ_ONLY` | `false` | `true` exposes only the stable `read` tools, whatever the other variables say. |
| `OFFICE_MCP_DISABLED_TOOLS` | unset | Comma-separated tool names to hide. |
| `OFFICE_MCP_EXPERIMENTAL` | unset | Comma-separated experimental tool names to register, or `all`. |
| `OFFICE_MCP_PERSONA_FILE` | `~/.config/office-mcp/persona.md` | The persona file. Setting it explicitly makes a missing file block sends. |
| `OFFICE_MCP_PERSONA_REQUIRED` | `false` | `true` blocks every send while the persona file is missing or invalid. |
| `OFFICE_MCP_SEND_ONLY_TO` | unset | Recipient allowlist: `user@domain`, `@domain`, `teams:self` and `teams:<thread id>` entries. Unset means no allowlist. |
| `OFFICE_MCP_SELF_CHAT_THREAD` | `48:notes` | Thread id that `teams:self` stands for. |
| `OFFICE_MCP_TEAMS_MAX_MEMBERS` | `25` | Group chats with more members need `group_ok=true` (2 to 1000). |
| `OFFICE_MCP_RESTORE_FRONTMOST` | `true` | Give focus back to the previous app after a tool brought Outlook or Teams forward. |
| `OFFICE_MCP_IDLE_SECONDS` | `1.5` | How long you must not have used the Mac before a tool brings an app forward (0 to 30, 0 turns the wait off). |
| `OFFICE_MCP_ACTION_TIMEOUT` | `12` | Seconds a single UI step waits for the app (1 to 120). |
| `OFFICE_MCP_LOCK_FILE` | `~/Library/Caches/office-mcp/ui.lock` | Cross-process lock shared by the server and the CLI. |
| `OFFICE_MCP_LOCK_TIMEOUT` | `20` | Seconds to wait for the lock before a call fails as busy (1 to 300). |
| `OFFICE_MCP_STALE_AFTER` | `10` | A write that waited longer than this for the lock refuses to act (1 to 300). |
| `OFFICE_MCP_DOWNLOAD_DIR` | unset | Absolute folder for `outlook_save_attachment`. The tool is registered only when it is set. |

A typical owner setup turns on everything except the invitation responses:

```sh
OFFICE_MCP_TOOLS=all
OFFICE_MCP_DISABLED_TOOLS=calendar_respond
OFFICE_MCP_PERSONA_FILE=~/.config/office-mcp/persona.md
OFFICE_MCP_PERSONA_REQUIRED=true
```

## Running

stdio, for clients that start the server themselves:

```sh
uv run --package office-mcp office-mcp
```

Streamable HTTP on `http://127.0.0.1:8769/mcp` (stateless, JSON responses):

```sh
uv run --package office-mcp office-mcp --http --port 8769
curl http://127.0.0.1:8769/health
```

To start HTTP mode at login, copy [`examples/local.office-mcp.plist`](examples/local.office-mcp.plist) to `~/Library/LaunchAgents/`, adjust the paths, and load it with `launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/local.office-mcp.plist`. Grant Accessibility to the Python binary that launchd starts.

### CLI

`office` runs the same tools and prints JSON. Every tool is a subcommand named after it with dashes (`office outlook-search budget --scope all`); required parameters are positional, list parameters repeat (`--to a@example.com --to b@example.com`), and booleans take `--flag` or `--no-flag`. The same exposure variables apply, so a tool that the server would not expose is refused. The old short names still work:

```sh
uv run --package office-mcp office status
uv run --package office-mcp office folders
uv run --package office-mcp office mail --folder Inbox --days 3 --view both
uv run --package office-mcp office read-mail <key>
uv run --package office-mcp office chats --unread-only
uv run --package office-mcp office read-chat "Jane Doe" --limit 10
uv run --package office-mcp office needs-reply --days 7
uv run --package office-mcp office office-check-text "Draft text" email
uv run --package office-mcp office persona-bootstrap --emails 40 --chats 15 --per-chat 20
```

`office read-mail` keeps the unread state by default; pass `--mark-read` (or `--no-keep-unread`) to leave it read.

## Connecting clients

### Claude Code

Over stdio:

```sh
claude mcp add office -- uv run --quiet --no-dev --project /path/to/mcp-suite/servers/office office-mcp
```

Over HTTP, with the server running:

```sh
claude mcp add --transport http --scope user office http://127.0.0.1:8769/mcp \
  --header "Authorization: Bearer your-token"
```

### Claude Desktop

Add this to `claude_desktop_config.json` and grant Accessibility to Claude.app:

```json
{
  "mcpServers": {
    "office": {
      "command": "uv",
      "args": ["run", "--quiet", "--no-dev", "--project", "/path/to/mcp-suite/servers/office", "office-mcp"],
      "env": { "OFFICE_MCP_TIMEZONE": "Europe/London" }
    }
  }
}
```

### Open WebUI

In Admin Settings, External Tools, add a server of type "MCP (Streamable HTTP)" with the URL `http://127.0.0.1:8769/mcp`. If Open WebUI runs in Docker, use `http://host.docker.internal:8769/mcp`, set `OFFICE_MCP_ALLOWED_HOSTS=host.docker.internal`, and set `OFFICE_MCP_TOKEN` and the matching bearer token in Open WebUI.

### From another machine over SSH

The server has to run on the Mac, but the client can live elsewhere. Two scripts in [`examples/`](examples) cover this. Both read `MAC_HOST` (required), plus optional `MAC_USER`, `MAC_REPO` (default `$HOME/mcp-suite`) and `MAC_UV` (default `$HOME/.local/bin/uv`). `ssh-tunnel.sh` also reads `LOCAL_PORT`, `REMOTE_PORT` (default 8769) and `START_REMOTE`.

- [`ssh-stdio.sh`](examples/ssh-stdio.sh) runs the server over SSH as a stdio MCP server.
- [`ssh-tunnel.sh`](examples/ssh-tunnel.sh) forwards a local port to the Mac and starts HTTP mode there for as long as the tunnel lives. [`office-mcp.service`](examples/office-mcp.service) is a systemd user unit that keeps the tunnel up, reading `~/.config/office-mcp/tunnel.env` (see [`tunnel.env.example`](examples/tunnel.env.example)).

A server started over SSH needs Accessibility granted to `/usr/libexec/sshd-keygen-wrapper` on the Mac.

## Live testing

The unit tests run anywhere against fakes. The real apps are tested by a self-only protocol on the Mac: every message goes to your own address and your own Teams chat, enforced by `OFFICE_MCP_SEND_ONLY_TO`. The steps, preconditions and pass rules are in [docs/live-testing.md](docs/live-testing.md). The harness in `tests/live/` runs only with `OFFICE_MCP_LIVE=1` and the `live` marker, and refuses unless `office_status.send_only_to` is exactly your address and your self chat.

## Security and limitations

- The Accessibility permission lets the granted process read and control the entire UI of the Mac. Grant it only to processes you trust.
- Granting Accessibility to `sshd-keygen-wrapper` gives that power to every SSH session on the Mac. Prefer a dedicated Mac or VM.
- The tools change what is on screen, and write tools bring the apps forward. They wait for you to be idle, but do not use the Mac interactively while a client is calling write tools.
- Sends go out at once. There is no draft step; the persona, allowlist and duplicate checks are the guard rails.
- Tool output (names, addresses, mail and chat content) goes to whatever model provider your MCP client uses. Treat it as confidential.
- Message content can contain prompt injection. Keep the write groups off for clients you do not fully control, or use `OFFICE_MCP_SEND_ONLY_TO`.
- HTTP mode has no authentication unless `OFFICE_MCP_TOKEN` is set. Keep it on loopback and use the SSH tunnel for remote access.
- This is UI scraping. A Microsoft update to Outlook or Teams can change labels or layout and break a tool until it is updated. Non-English UIs and other date formats are not supported.

## Development

From `servers/office`:

```sh
uv run ruff check .
uv run ruff format --check .
uv run mypy src tests
uv run pytest
uv run python -I scripts/no_comments.py src tests scripts
uv run python -I scripts/contract_check.py
```

The tests run on any OS. pyobjc is imported only when the server first touches the apps, and every feature is tested against fake accessibility trees, gateways and scopes. Live tests are excluded by default (`-m 'not live'`).

Layout:

```
src/office_mcp/
  server.py          MCP server, transports, registration
  cli.py             office CLI generated from the tool specs
  container.py       wires settings, drivers, services and tool specs
  tools.py           office_status
  bootstrap.py       sources for persona-bootstrap
  descriptions.py    server instructions
  features/outlook/  folders, list, reading, search, people, organize, compose
  features/calendar/ Teams calendar events and responses
  features/teams/    chats, messages, activity, search, channels, writing
  features/persona/  persona file, rules, guard and bootstrap
  shared/            settings, registry, accessibility core, focus, menus, runner, text
```

## Troubleshooting

- **"Outlook is not working right now. Outlook is offline..."**: press Go Online in the yellow banner or sign in again, then retry.
- **"Outlook is not working right now. Outlook's view is stuck on..."**: quit and reopen Outlook.
- **`accessibility: false`** in `office status`: grant Accessibility to the process that starts the server and restart it.
- **BusyError "another office action is still running"**: another call (or the CLI) holds the lock. Wait and retry. "the owner is using the Mac" means you were typing or moving the mouse; the tool waits up to 10 seconds for `OFFICE_MCP_IDLE_SECONDS` of idle time.
- **ActivationError "could not bring ... to the front"**: another app or a dialog kept focus. Close the dialog and retry.
- **SessionUnavailableError**: the screen is locked or a password field has secure input on. Unlock the Mac or leave the password field.
- **MenuItemDisabledError**: the app was not frontmost or nothing was selected, so the menu item was off. Retry when the Mac is idle.
- **PersonaViolationError "Not sent: the text breaks N persona rules"**: fix each listed problem and call again, or test with `office_check_text` first. "The persona file is missing" or "invalid" means sends stay blocked until the file is fixed; `office_persona` shows the error.
- **RecipientNotAllowedError**: the target is not in `OFFICE_MCP_SEND_ONLY_TO`. Nothing was sent.
- **SendRefusedError**: Outlook or Teams showed an alert after Send, or Teams sent part of the text early. The alert text is in the message; nothing more was sent.
- **ComposeBusyError**: a compose window, reply or chat box already holds text (perhaps your own unfinished draft). Finish or discard it in the app first.
- **DuplicateSendError**: the same text went out recently. Pass `allow_duplicate=true` only when it really must go out again.
- **"office-mcp reads Outlook and Teams through the macOS accessibility API..."**: the server runs on a machine that is not a Mac, or pyobjc is missing.
- **Wrong Teams message timestamps**: set `OFFICE_MCP_TIMEZONE` to your zone.
- **421 "Invalid Host header"** over HTTP: add the host name your client uses to `OFFICE_MCP_ALLOWED_HOSTS`.
- **401 unauthorized**: the client is not sending `Authorization: Bearer <token>` matching `OFFICE_MCP_TOKEN`.
