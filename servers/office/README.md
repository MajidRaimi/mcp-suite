# office-mcp

An MCP server that reads Microsoft Outlook mail and Microsoft Teams chats from the desktop apps that are already signed in on a Mac. It works through the macOS Accessibility API, so it sees what the apps show and needs no Microsoft Graph app registration, API keys or tenant configuration.

Every tool returns data and nothing is sent, deleted or moved. The tools do drive the visible UI, though: they switch folders, select conversations and open chats on screen, and opening a Teams chat marks it read. Read [Security and limitations](#security-and-limitations) before you run it.

## Tools

| Tool | Parameters (defaults) | Returns | Side effects |
|---|---|---|---|
| `office_status` | none | Whether the Accessibility permission is granted and whether Outlook and Teams are running | None |
| `outlook_folders` | none | Mail folders with unread counts | None |
| `outlook_conversations` | `folder` ("Inbox"), `unread_only` (false), `days` (none), `limit` (50, 0 means all), `load_more` (0) | Conversations newest first, each with a stable `key`, flags, unread and message counts, title, time label, best-effort ISO date, meeting details and a preview | Switches Outlook to `folder`. `days` clicks "Load more conversations" until the range is covered (at most 20 rounds); `load_more` clicks it that many times |
| `outlook_read` | `ref`, `folder` (current folder), `keep_unread` (true), `max_body` (20000) | The latest email in the conversation: subject, tags, sender, recipients by field with emails, collapsed recipient counts, sent time, notes and the body including quoted history | Selects the conversation (and switches to `folder` when given). Outlook marks it read when selected; with `keep_unread=true` an email that was unread is marked unread again, with `keep_unread=false` it stays read |
| `teams_chats` | `unread_only` (false), `section` (none), `limit` (0, means all) | Chats from the chat list, top to bottom, with section, kind, unread and muted flags, presence, last message preview, sender and date | Switches Teams to the Chat view if another view is showing |
| `teams_read` | `chat`, `limit` (30), `older` (0) | The chat and its latest messages with author, ownership, ISO timestamp and text | Opens the chat on screen, so Teams marks it read. This is not reverted. `older` scrolls up that many times to load older messages |
| `teams_needs_reply` | `days` (7), `include_meetings` (false) | Chats whose last message is not from you, not muted and within `days`, unread first | Reads the chat list only, so nothing is marked read |

`ref` for `outlook_read` is a `key` from `outlook_conversations` (best), a row index, or words from the title. `chat` for `teams_read` is a chat name, part of one, or a thread id. Ambiguous names return an error that lists the candidates.

All tools share one lock, so calls run one at a time. A single UI step waits up to 12 seconds for the app to respond (Teams also waits up to 12 seconds for its window to settle on first use, and an app that is not open gets 15 seconds to launch). A call that loads many pages can take much longer than that.

Every tool carries MCP annotations with `destructive_hint=false` and `open_world_hint=false`. `teams_read` and `outlook_read` can change unread state, so they have `read_only_hint=false` and `idempotent_hint=false`. Every other tool has `read_only_hint=true` and `idempotent_hint=true`.

## How it works

The server uses [pyobjc](https://pyobjc.readthedocs.io) to call the macOS Accessibility API (`AXUIElement`) against the running Outlook and Teams processes. It reads the folder pane, message list, reading pane and the Teams chat list and message pane from the accessibility tree, and presses rows and buttons the same way assistive technology does. Because it uses the desktop apps you are already signed in to, it needs no Microsoft Graph app registration and no credentials of its own. It only sees what your account can already see in those apps.

## Requirements

- macOS with a logged-in GUI session (the apps must be on screen, not just running in the background of a locked or headless session).
- The new Outlook for Mac (bundle id `com.microsoft.Outlook`), signed in.
- The new Microsoft Teams (bundle id `com.microsoft.teams2`), signed in.
- English UI in both apps. The parser matches English labels such as "Load more conversations", "Mark Unread" and "Unread message".
- Date formats: the Outlook message list must show dates as D/M/YYYY with 12-hour times (for example `18/10/2026` and `6:53 AM`). The Teams chat list shows dates as M/D or M/D/YYYY (for example `10/19`).
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

   All three values (`accessibility`, `outlook_running`, `teams_running`) should be `true`.

## Configuration

The server reads these variables from its environment. `.env.example` lists them with their defaults. The server does not load a `.env` file by itself; export the variables, put them in your client's `env` block, or run with `uv run --env-file servers/office/.env ...`.

| Variable | Default | Meaning |
|---|---|---|
| `OFFICE_MCP_TRANSPORT` | `stdio` | `stdio` or `streamable-http`. `--transport` and `--http` override it. |
| `OFFICE_MCP_HOST` | `127.0.0.1` | HTTP bind address. `--host` overrides it. |
| `OFFICE_MCP_PORT` | `8769` | HTTP port. `--port` overrides it. |
| `OFFICE_MCP_TIMEZONE` | system zone, else UTC | IANA zone used for Teams message timestamps. List labels such as `10:42`, `Yesterday` or `Mon` (and so the `days` filters and `teams_needs_reply`) are read against the Mac's own date, because the apps render them in the Mac's system zone. |
| `OFFICE_MCP_TOKEN` | unset | When set, every HTTP request (including `/health`) must send `Authorization: Bearer <token>`, otherwise it gets a 401. Ignored for stdio. |
| `OFFICE_MCP_ALLOWED_HOSTS` | unset | Comma-separated `host` or `host:port` values accepted in the `Host` and `Origin` headers, in addition to loopback. A bare host also allows any port. Needed for Open WebUI in Docker (`host.docker.internal`) or a reverse proxy that keeps the original `Host`. When unset, the MCP SDK default applies. |
| `OFFICE_MCP_OWNER_NAME` | unset | Name used in the server instructions and tool descriptions. Falls back to "the user". |

Values are trimmed, and a blank value counts as unset. Invalid values (a port that is not a number, an unknown zone or transport, an allowed host that contains `/` or whitespace) stop the server at startup with a clear message.

## Running

stdio, for clients that start the server themselves:

```sh
uv run --package office-mcp office-mcp
```

Streamable HTTP on `http://127.0.0.1:8769/mcp` (stateless, JSON responses):

```sh
uv run --package office-mcp office-mcp --transport streamable-http
uv run --package office-mcp office-mcp --http --port 8769
curl http://127.0.0.1:8769/health
```

To start HTTP mode at login, copy [`examples/local.office-mcp.plist`](examples/local.office-mcp.plist) to `~/Library/LaunchAgents/`, adjust the paths, and load it with `launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/local.office-mcp.plist`. Grant Accessibility to the Python binary that launchd starts.

### Debug CLI

`office` calls the same code as the tools and prints JSON:

```sh
uv run --package office-mcp office status
uv run --package office-mcp office folders
uv run --package office-mcp office mail --folder Inbox --days 3
uv run --package office-mcp office read-mail <key>
uv run --package office-mcp office chats --unread-only
uv run --package office-mcp office read-chat "Alice" --limit 10
uv run --package office-mcp office needs-reply --days 7
```

`office read-mail` keeps the unread state by default; pass `--mark-read` to leave it read.

## Connecting clients

### Claude Code

Over stdio:

```sh
claude mcp add office -- uv run --quiet --no-dev --project /path/to/mcp-suite/servers/office office-mcp
```

Over HTTP, with the server running:

```sh
claude mcp add --transport http --scope user office http://127.0.0.1:8769/mcp
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

Use the absolute path to `uv` if Claude Desktop cannot find it.

### Open WebUI

In Admin Settings, External Tools, add a server of type "MCP (Streamable HTTP)" with the URL `http://127.0.0.1:8769/mcp`. If Open WebUI runs in Docker, use `http://host.docker.internal:8769/mcp`, set `OFFICE_MCP_ALLOWED_HOSTS=host.docker.internal`, and set `OFFICE_MCP_TOKEN` and the matching bearer token in Open WebUI. Keep the server on loopback rather than binding it to a public interface.

### From another machine over SSH

The server has to run on the Mac, but the client can live elsewhere, for example on a Linux box. Two scripts in [`examples/`](examples) cover this. Both read `MAC_HOST` (required), plus optional `MAC_USER` (default: your local user), `MAC_REPO` (default `$HOME/mcp-suite`) and `MAC_UV` (default `$HOME/.local/bin/uv`). `ssh-tunnel.sh` also reads `LOCAL_PORT` and `REMOTE_PORT` (default 8769) and `START_REMOTE`. Both expect key-based SSH to the Mac and a clone of this repo there.

- [`ssh-stdio.sh`](examples/ssh-stdio.sh) runs the server over SSH as a stdio MCP server. Point a stdio client at it:

  ```sh
  claude mcp add office --env MAC_HOST=mac.example.internal --env MAC_USER=you -- /path/to/mcp-suite/servers/office/examples/ssh-stdio.sh
  ```

- [`ssh-tunnel.sh`](examples/ssh-tunnel.sh) forwards `127.0.0.1:$LOCAL_PORT` on the local machine to `$REMOTE_PORT` on the Mac and starts HTTP mode there for as long as the tunnel lives. Set `START_REMOTE=0` when the Mac already runs the server (for example from the launchd plist) to forward the port only. [`office-mcp.service`](examples/office-mcp.service) is a systemd user unit that keeps the tunnel up; it reads its variables from `~/.config/office-mcp/tunnel.env` (see [`tunnel.env.example`](examples/tunnel.env.example)):

  ```sh
  mkdir -p ~/.config/office-mcp ~/.config/systemd/user
  cp servers/office/examples/tunnel.env.example ~/.config/office-mcp/tunnel.env
  cp servers/office/examples/office-mcp.service ~/.config/systemd/user/
  systemctl --user daemon-reload
  systemctl --user enable --now office-mcp.service
  ```

  Adjust `ExecStart` in the unit if your clone is not at `~/mcp-suite`.

A server started over SSH needs Accessibility granted to `/usr/libexec/sshd-keygen-wrapper` on the Mac.

## Security and limitations

- The Accessibility permission lets the granted process read and control the entire UI of the Mac, not just Outlook and Teams. Grant it only to processes you trust.
- Granting Accessibility to `sshd-keygen-wrapper` gives that power to every SSH session on the Mac, for every user who can log in. Prefer a dedicated Mac or VM for this.
- The tools change what is on screen: Outlook switches folders and selects conversations, Teams opens chats. Do not use the Mac interactively while a client is calling tools.
- `teams_read` marks the chat read, and that is not reverted. `outlook_read` restores the unread state only when `keep_unread` is true.
- Tool output (names, email addresses, mail and chat content) goes to whatever model provider your MCP client uses. Treat it as confidential.
- Message content can contain prompt injection. The server is read-only, but other tools in the same client may not be.
- HTTP mode has no authentication unless `OFFICE_MCP_TOKEN` is set. Keep it on loopback and use the SSH tunnel for remote access.
- This is UI scraping. A Microsoft update to Outlook or Teams can change labels or layout and break parsing until the parser is updated. Non-English UIs and other date formats are not supported.
- Calls are serialized by a global lock and can take several seconds, longer when loading older pages.

## Development

From the repository root:

```sh
uv sync --all-packages
uv run --package office-mcp ruff check servers/office
uv run ruff format --check servers/office
uv run --package office-mcp mypy --config-file servers/office/pyproject.toml servers/office/src servers/office/tests
uv run --package office-mcp pytest servers/office
```

Or from `servers/office`, run `uv run pytest` and `uv run mypy src tests`.

The tests run on any OS. pyobjc is imported only when the server first touches the apps, and the Outlook and Teams services are tested against fake gateways that feed fixed accessibility trees, so they cover parsing, folder and chat resolution, waiting and timeouts, settings, tool annotations and the HTTP token and host checks.

Layout:

```
src/office_mcp/
  server.py          MCP server, transports, CLI flags
  cli.py             office debug CLI
  tools.py           tool functions shared by the server and the CLI
  container.py       wires services to the accessibility gateways
  descriptions.py    instructions and tool descriptions
  features/outlook/  folders, message list and reading pane
  features/teams/    chat list, messages and needs-reply
  shared/            settings, HTTP middleware, accessibility helpers, clock, text
```

## Troubleshooting

- **`accessibility: false`** in `office status`: grant Accessibility to the process that starts the server (see [Setup](#setup)) and restart it. Remove and re-add the entry if it was granted before an update.
- **"office-mcp reads Outlook and Teams through the macOS accessibility API..."**: the server is running on a machine that is not a Mac, or pyobjc is missing. Run it on the Mac and use the SSH examples to reach it.
- **"Microsoft Outlook has no matching window"**: the main Outlook window title contains a "•" separator. Make sure the main mail window is open, not only a message or calendar window.
- **"Teams did not open the chat in time"** or **"Outlook did not show the conversation"**: the app was slow or a dialog was in the way. Bring the app to the front and retry.
- **Wrong Teams message timestamps**: set `OFFICE_MCP_TIMEZONE` to your zone.
- **"today" or the `days` filters off by a day**: check the Mac's own time zone in System Settings, since list labels follow it.
- **421 "Invalid Host header"** over HTTP: add the host name your client uses to `OFFICE_MCP_ALLOWED_HOSTS`.
- **401 unauthorized**: the client is not sending `Authorization: Bearer <token>` matching `OFFICE_MCP_TOKEN`.
- **Empty lists or missing dates**: check that both apps use an English UI and the date formats listed in [Requirements](#requirements).
