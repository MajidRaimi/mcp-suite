# whatsapp-mcp

An MCP server that lets an AI client read your WhatsApp and send replies you have confirmed. It talks to a self-hosted [WAHA](https://waha.devlike.pro) instance that is paired with your phone.

Platform: Linux or macOS (on Windows, run it inside WSL), with Python 3.13 and Docker. Access: read, plus write (sends real messages after a two-step draft and send).

## What it does

| Tool | Purpose | Key parameters (default, limits) | Side effects |
|---|---|---|---|
| `whatsapp_needs_reply` | Lists direct chats (optionally groups) whose last message is from the other person, oldest wait first, with wait time and unanswered count. | `limit` (20, 1 to 50), `include_groups` (false), `max_age_days` (7, 1 to 60) | None. Reads up to 300 chats plus one message page per candidate. |
| `whatsapp_recent_chats` | Lists the most recently active chats with their last message. | `limit` (20, 1 to 100) | None. |
| `whatsapp_read_chat` | Reads the latest messages of one chat, oldest first. `chat` can be a chat id, a contact or group name, or a phone number. | `chat`, `limit` (40, 1 to 200) | None. Ambiguous names return the candidates instead of guessing. |
| `whatsapp_search` | Searches recent messages and chat names for a word or phrase. | `query` (at least 2 characters), `days` (14, 1 to 60), `chats` (40, 1 to 80) | None. Returns at most 60 hits and makes up to 81 WAHA calls. |
| `whatsapp_draft_reply` | Prepares a message for one chat without sending it and returns a `draft_id`. | `chat`, `text` (1 to 4000 characters), `attachments` (absolute paths inside `WHATSAPP_MCP_ATTACHMENTS_DIR`, up to 10 files of 64 MB each) | Writes the draft to the local drafts file. Drafts expire after `WHATSAPP_MCP_DRAFT_TTL_MINUTES`. |
| `whatsapp_send_draft` | Sends a draft. Only for when the user has explicitly said to send it. | `draft_id` | Sends real WhatsApp messages. Not idempotent. The draft is consumed on success. |
| `whatsapp_status` | Shows the WAHA session state (for example `WORKING`). | none | None. |

Every tool carries MCP tool annotations: the five reading tools are read-only and idempotent, `whatsapp_draft_reply` writes local state only, and `whatsapp_send_draft` is open-world and not idempotent.

## How it works

The server calls the WAHA REST API (`/api/{session}/chats/overview`, `/api/{session}/chats/{id}/messages`, `/api/sendText`, `/api/sendFile`, `/api/sessions/{session}`) with the `X-Api-Key` header. WAHA keeps the linked-device connection to WhatsApp. Times are formatted in `WHATSAPP_MCP_TIMEZONE`.

Sending is two steps. `whatsapp_draft_reply` stores a draft with a short id and returns exactly what would be sent. `whatsapp_send_draft` sends that stored draft, so the model cannot change the text or recipient between the preview and the send. Attachments are hashed when drafted and checked again before sending: a file that changed, moved outside the allowed directory, or became a symlink that points outside it is refused.

## Requirements

- Python 3.13 or newer and [uv](https://docs.astral.sh/uv/).
- Docker with the compose plugin, to run WAHA.
- A phone with WhatsApp to pair the WAHA session.
- File attachments use WAHA's `sendFile` endpoint, which may require WAHA Plus depending on the WAHA edition and engine. Text messages work on the free image.

## Setup

1. Start WAHA from the repository root:

   ```sh
   cd servers/whatsapp/deploy/waha
   cp .env.example .env
   chmod 600 .env
   ```

   Fill in `.env`. Generate the key with `openssl rand -hex 32` and put it in both `WAHA_API_KEY` and `WAHA_API_KEY_PLAIN` (WAHA also accepts `WAHA_API_KEY=sha512:<hex digest of the key>`). Use different strong passwords for the dashboard and Swagger. Then:

   ```sh
   docker compose up -d
   ```

   WAHA listens on `127.0.0.1:3500` only. The session data lives in `./sessions`.

2. Pair your phone. Open `http://127.0.0.1:3500/dashboard`, log in with the dashboard credentials, connect it to `http://127.0.0.1:3500` with your API key, and start the session named `default` (WAHA's free image runs one session with that name). Scan the QR code from WhatsApp on your phone under Linked devices. Wait until the session status is `WORKING`.

3. Install the server:

   ```sh
   uv sync --package whatsapp-mcp
   ```

## Configuration

All settings are environment variables. `servers/whatsapp/.env.example` lists them with safe defaults. The server does not load a .env file by itself; export the variables, put them in your client's env block, or run with `uv run --env-file servers/whatsapp/.env --package whatsapp-mcp whatsapp-mcp`.

| Variable | Required | Default | Meaning |
|---|---|---|---|
| `WAHA_API_KEY_PLAIN` | yes | none | The plain WAHA API key sent as `X-Api-Key`. |
| `WAHA_URL` | no | `http://127.0.0.1:3500` | Base URL of WAHA. |
| `WAHA_SESSION` | no | `default` | WAHA session name. |
| `WHATSAPP_MCP_TRANSPORT` | no | `streamable-http` | `stdio` or `streamable-http`. The `--transport` flag overrides it. |
| `WHATSAPP_MCP_HOST` | no | `127.0.0.1` | HTTP bind address. The `--host` flag overrides it. |
| `WHATSAPP_MCP_PORT` | no | `8768` | HTTP port, 1 to 65535. The `--port` flag overrides it. |
| `WHATSAPP_MCP_TOKEN` | no | unset | When set, every HTTP request (including `/health`) must send `Authorization: Bearer <token>`, otherwise it gets a 401. Ignored for stdio. |
| `WHATSAPP_MCP_ALLOWED_HOSTS` | no | unset | Comma-separated `host` or `host:port` values accepted in the `Host` and `Origin` headers, in addition to loopback. Needed for Open WebUI in Docker (`host.docker.internal`) or a reverse proxy. When unset, the MCP SDK default applies: loopback only when bound to loopback. |
| `WHATSAPP_MCP_TIMEZONE` | no | system zone, else UTC | IANA zone used to format message times, for example `Europe/Berlin`. |
| `WHATSAPP_MCP_OWNER_NAME` | no | unset | Name used in tool descriptions and instructions. Without it they say "the user". |
| `WHATSAPP_MCP_ATTACHMENTS_DIR` | no | unset | Absolute path of the only directory that attachments may come from (symlinks are resolved). When unset, attachments are refused. |
| `WHATSAPP_MCP_DRAFTS` | no | `~/.local/state/whatsapp-mcp/drafts.json` | Where pending drafts are stored (mode 600). |
| `WHATSAPP_MCP_DRAFT_TTL_MINUTES` | no | `30` | Draft lifetime in minutes, 1 to 10080. |

Invalid values stop the server at startup with a message that names the variable.

## Running

HTTP (the default, for clients that connect by URL), from the repository root. The WAHA `.env` already holds `WAHA_API_KEY_PLAIN`, so the key stays out of your shell history:

```sh
uv run --env-file servers/whatsapp/deploy/waha/.env --package whatsapp-mcp whatsapp-mcp
```

Add `--host` and `--port` to override the environment, for example `--host 127.0.0.1 --port 8768`. The MCP endpoint is `http://127.0.0.1:8768/mcp` (stateless, JSON responses). Check it in a second terminal:

```sh
curl -s http://127.0.0.1:8768/health
```

It answers `{"status":"WORKING"}` with 200 when the WAHA session is connected, and the session status (or `UNREACHABLE`) with 503 otherwise. Add `-H "Authorization: Bearer $WHATSAPP_MCP_TOKEN"` when a token is set.

stdio (normally started by your client, see below; run by hand it waits silently on stdin):

```sh
uv run --env-file servers/whatsapp/deploy/waha/.env --package whatsapp-mcp whatsapp-mcp --transport stdio
```

## Connecting clients

Claude Code over HTTP:

```sh
claude mcp add --transport http --scope user whatsapp http://127.0.0.1:8768/mcp
claude mcp add --transport http --scope user whatsapp http://127.0.0.1:8768/mcp \
  --header "Authorization: Bearer your-token"
```

Claude Code over stdio:

```sh
claude mcp add whatsapp --env WAHA_API_KEY_PLAIN=your-key -- \
  uv run --quiet --project /ABS/PATH/mcp-suite/servers/whatsapp whatsapp-mcp --transport stdio
```

Claude Desktop (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "whatsapp": {
      "command": "/absolute/path/to/uv",
      "args": [
        "run", "--quiet", "--project", "/ABS/PATH/mcp-suite/servers/whatsapp",
        "whatsapp-mcp", "--transport", "stdio"
      ],
      "env": {
        "WAHA_API_KEY_PLAIN": "your-key",
        "WHATSAPP_MCP_ATTACHMENTS_DIR": "/ABS/PATH/whatsapp-outbox"
      }
    }
  }
}
```

Use the absolute path of uv (`which uv`); Claude Desktop does not always inherit your shell PATH.

Open WebUI: Admin Settings, External Tools, add a connection of type "MCP (Streamable HTTP)" with the URL `http://127.0.0.1:8768/mcp` and, when a token is set, Bearer authentication with that token.

If Open WebUI runs in Docker, `127.0.0.1` inside the container is the container itself. Either run Open WebUI with host networking, or bind the server to the Docker bridge address and allow the Docker host name:

```sh
WHATSAPP_MCP_HOST=172.17.0.1
WHATSAPP_MCP_ALLOWED_HOSTS=host.docker.internal
WHATSAPP_MCP_TOKEN=a-long-random-token
```

and use `http://host.docker.internal:8768/mcp` in Open WebUI (start the container with `--add-host=host.docker.internal:host-gateway` on Linux). Never bind to `0.0.0.0` on a public interface.

## Persistent deployment

The included user unit runs the server under systemd for one user. From the repository root:

```sh
uv tool install ./servers/whatsapp
mkdir -p ~/.config/whatsapp-mcp
cp servers/whatsapp/.env.example ~/.config/whatsapp-mcp/env
chmod 600 ~/.config/whatsapp-mcp/env
mkdir -p ~/.config/systemd/user
cp servers/whatsapp/deploy/systemd/whatsapp-mcp.service ~/.config/systemd/user/
```

Edit `~/.config/whatsapp-mcp/env` (at least `WAHA_API_KEY_PLAIN`). Then start the service:

```sh
systemctl --user daemon-reload
systemctl --user enable --now whatsapp-mcp
loginctl enable-linger "$USER"
```

`loginctl enable-linger` keeps the service running when you are logged out. `uv tool install` does not use the workspace lock file, so the dependency ranges in `pyproject.toml` (`mcp>=2.3,<3`) are what pin that install.

## Security and limitations

- WAHA is an unofficial WhatsApp client. Using it can break WhatsApp's Terms of Service and can get the number banned. Use it at your own risk, ideally with a number you can afford to lose.
- Message content is untrusted text written by other people and goes straight into the model. A message can try to instruct the model (prompt injection). Everything the tools return is also sent to whichever model provider your client uses.
- The "only send when the user says so" rule is enforced by prompt text, not by the server. Keep tool approval turned on in your client for `whatsapp_send_draft`, and read the draft before approving.
- Sends are not idempotent. Retrying a send that timed out can deliver the message twice.
- A draft with several attachments is sent one file at a time. If a later file fails, the earlier files are already delivered, the draft is not restored, and the error says how many items went out. Do not resend the whole draft.
- Attachments are read from the server host's filesystem. Keep `WHATSAPP_MCP_ATTACHMENTS_DIR` pointed at a dedicated directory that holds only files you are willing to send.
- The HTTP endpoint has no authentication unless `WHATSAPP_MCP_TOKEN` is set. Keep it on loopback, and use a token plus `WHATSAPP_MCP_ALLOWED_HOSTS` whenever it is reachable from anything else.

## Development

From the repository root:

```sh
uv sync --all-packages
uv run --package whatsapp-mcp ruff check servers/whatsapp
uv run ruff format --check servers/whatsapp
uv run --package whatsapp-mcp mypy --config-file servers/whatsapp/pyproject.toml servers/whatsapp/src servers/whatsapp/tests
uv run --package whatsapp-mcp pytest servers/whatsapp
```

Or from `servers/whatsapp`: `uv run pytest` and `uv run mypy src tests`. The tests use `httpx.MockTransport` and Starlette's test client, so they need no network, no WAHA and no WhatsApp account.

## Troubleshooting

- `whatsapp_status` or `/health` shows `SCAN_QR_CODE`: the session is not paired yet, or WhatsApp logged it out. Open the WAHA dashboard and scan the QR code again.
- `/health` returns 503 with `UNREACHABLE`: the server cannot reach WAHA or WAHA rejected the key. Check `WAHA_URL` (the bundled compose file publishes port 3500, not WAHA's stock 3000), `WAHA_API_KEY_PLAIN` and `WAHA_SESSION`, and run `docker compose -f servers/whatsapp/deploy/waha/compose.yml logs waha` from the repository root.
- `/mcp` returns 421 (Invalid Host header) or 403 (Invalid Origin header): the request came through a host name the MCP SDK does not trust, for example `host.docker.internal` or a reverse proxy. Add that host to `WHATSAPP_MCP_ALLOWED_HOSTS`.
- Every request returns 401: `WHATSAPP_MCP_TOKEN` is set and the client is not sending `Authorization: Bearer <token>`.
- `file attachments are disabled`: set `WHATSAPP_MCP_ATTACHMENTS_DIR` and restart the server.
- Message times look off by some hours: set `WHATSAPP_MCP_TIMEZONE`.
