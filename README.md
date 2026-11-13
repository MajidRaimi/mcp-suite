# mcp-suite

Self-hostable [MCP](https://modelcontextprotocol.io) servers that connect your own accounts and desktop apps to any MCP client: Claude Code, Claude Desktop, Open WebUI, or anything else that speaks stdio or streamable HTTP. Each server is independent, runs on your machine, keeps your data local-first, and binds to loopback by default. Clone the repo, run one command, and point your client at the server you need.

## Servers

| Server | Platform | What it does | Access | Transports | Default port |
|---|---|---|---|---|---|
| [whatsapp](servers/whatsapp) | Any OS with Python 3.13 and Docker | Lists chats waiting on a reply, reads and searches chats, drafts replies and sends them after explicit confirmation, through a self-hosted [WAHA](https://waha.devlike.pro) | Read + confirmed send | stdio, streamable HTTP | 8768 |
| [office](servers/office) | macOS only | Reads Outlook folders and mail, lists and reads Teams chats, and finds chats awaiting a reply, through the macOS Accessibility API | Read-only data, but it drives the visible UI | stdio, streamable HTTP | 8769 |

The next free port is **8770**. Claim it in this table when you add a server.

## Requirements

- [git](https://git-scm.com)
- [uv](https://docs.astral.sh/uv/getting-started/installation/)
- Python 3.13. You do not need to install it yourself: uv downloads it on first sync.

Per server:

- **whatsapp**: Docker (for WAHA) and a phone with WhatsApp to pair.
- **office**: a Mac with the new Outlook and new Teams apps signed in, and the Accessibility permission for the process that runs the server.

## Quickstart

```sh
git clone https://github.com/MajidRaimi/mcp-suite
cd mcp-suite
uv sync --all-packages
```

That installs every server into one virtual environment at `.venv`. Then follow the block for the server you want. Each server README has the full setup, every setting and troubleshooting.

### whatsapp

```sh
cp servers/whatsapp/deploy/waha/.env.example servers/whatsapp/deploy/waha/.env
chmod 600 servers/whatsapp/deploy/waha/.env
```

Edit that file: generate a key with `openssl rand -hex 32`, put it in both `WAHA_API_KEY` and `WAHA_API_KEY_PLAIN`, and set two different strong passwords for the dashboard and Swagger. Then start WAHA:

```sh
docker compose -f servers/whatsapp/deploy/waha/compose.yml up -d
```

Pair your phone in the WAHA dashboard, then start the server. It reads `WAHA_API_KEY_PLAIN` from the same file, so the key never lands in your shell history:

```sh
uv run --env-file servers/whatsapp/deploy/waha/.env --package whatsapp-mcp whatsapp-mcp
```

In a second terminal:

```sh
curl http://127.0.0.1:8768/health
```

Full guide: [servers/whatsapp/README.md](servers/whatsapp/README.md).

### office

On the Mac, grant Accessibility to your terminal (System Settings, Privacy and Security, Accessibility), then check the setup:

```sh
uv run --package office-mcp office status
```

`office status` should report every check as true. Over stdio your client starts the server itself (see [Connecting clients](#connecting-clients)), so there is nothing to run by hand. For clients that connect by URL, start it over HTTP:

```sh
uv run --package office-mcp office-mcp --transport streamable-http
```

Full guide: [servers/office/README.md](servers/office/README.md).

## Connecting clients

Every HTTP server listens on `http://127.0.0.1:<port>/mcp`. Every server also runs over stdio, where the client starts it as a subprocess.

### Claude Code

Over HTTP, with the server already running:

```sh
claude mcp add --transport http --scope user whatsapp http://127.0.0.1:8768/mcp
```

When the server has a token (`WHATSAPP_MCP_TOKEN`):

```sh
claude mcp add --transport http --scope user whatsapp http://127.0.0.1:8768/mcp \
  --header "Authorization: Bearer your-token"
```

Over stdio, where Claude Code starts the server for you:

```sh
claude mcp add office -- uv run --project /path/to/mcp-suite/servers/office office-mcp
```

### Claude Desktop

Edit `claude_desktop_config.json` (Settings, Developer, Edit Config). Use the absolute path of `uv` (`which uv`) because desktop apps do not always inherit your shell `PATH`.

```json
{
  "mcpServers": {
    "office": {
      "command": "/absolute/path/to/uv",
      "args": ["run", "--project", "/path/to/mcp-suite/servers/office", "office-mcp"],
      "env": {
        "OFFICE_MCP_TIMEZONE": "Europe/London"
      }
    },
    "whatsapp": {
      "command": "/absolute/path/to/uv",
      "args": ["run", "--project", "/path/to/mcp-suite/servers/whatsapp", "whatsapp-mcp"],
      "env": {
        "WHATSAPP_MCP_TRANSPORT": "stdio",
        "WAHA_API_KEY_PLAIN": "your-key"
      }
    }
  }
}
```

For office over stdio, macOS asks for the Accessibility permission on behalf of Claude Desktop itself, so grant it to Claude.

### Open WebUI

1. Start the server with `--transport streamable-http` (whatsapp does this by default).
2. In Open WebUI, open Admin Settings, External Tools, add a connection, and choose **MCP (Streamable HTTP)**.
3. URL: `http://127.0.0.1:8768/mcp` if Open WebUI runs directly on the host or in a container with host networking.
4. If Open WebUI runs in Docker with its own network, use `http://host.docker.internal:8768/mcp` and set `WHATSAPP_MCP_ALLOWED_HOSTS=host.docker.internal`. Without it, the MCP SDK rejects the request with `421 Invalid Host header`. On Linux, start the container with `--add-host=host.docker.internal:host-gateway` and bind the server to the Docker bridge address (`WHATSAPP_MCP_HOST=172.17.0.1`), never to `0.0.0.0`.
5. Whenever the server is bound to the Docker bridge address (or any address other than loopback), set `WHATSAPP_MCP_TOKEN` and `WHATSAPP_MCP_ALLOWED_HOSTS`, then set the connection's auth to Bearer and paste the token. If the server stays on loopback and has a token, do the same.

### Any other client

Use the streamable HTTP endpoint `/mcp`. The servers run stateless (`stateless_http=True`) and answer with plain JSON (`json_response=True`), so no session id or SSE support is needed. Send `Authorization: Bearer <token>` when a token is configured.

## Common configuration

Every server reads the same set of variables, with its own prefix (`WHATSAPP_MCP_`, `OFFICE_MCP_`, and so on). Server-specific variables are listed in each server README and `.env.example`.

| Variable | Default | Meaning |
|---|---|---|
| `<NAME>_MCP_TRANSPORT` | whatsapp: `streamable-http`, office: `stdio` | `stdio` or `streamable-http`. The `--transport` flag overrides it. |
| `<NAME>_MCP_HOST` | `127.0.0.1` | HTTP bind address. Keep it on loopback unless you know why not. |
| `<NAME>_MCP_PORT` | the server's port above | HTTP port. Must be an integer between 1 and 65535. |
| `<NAME>_MCP_TOKEN` | empty (no auth) | When set, every HTTP request, including `/health`, needs `Authorization: Bearer <token>` or gets a `401`. Ignored over stdio. |
| `<NAME>_MCP_ALLOWED_HOSTS` | empty | Comma-separated `host[:port]` values to accept in the `Host` and `Origin` headers, on top of loopback. Needed behind a reverse proxy or for `host.docker.internal`. Setting it also turns the check on when the server is bound to a non-loopback address. |
| `<NAME>_MCP_TIMEZONE` | system zone, else UTC | IANA zone such as `America/New_York`, used for every time the server shows. |
| `<NAME>_MCP_OWNER_NAME` | `the user` | Name used in the server instructions and tool descriptions, so the model knows whose account it is. |

Generate a token with `openssl rand -hex 32` or `uv run python -c "import secrets; print(secrets.token_hex(32))"`.

## Security model

Read this before you expose anything. The full policy and how to report a vulnerability are in [SECURITY.md](SECURITY.md).

- **No auth by default, loopback only.** Out of the box, servers bind to `127.0.0.1` and accept any local request. For remote access, set `<NAME>_MCP_TOKEN` and reach the server through an SSH tunnel, a private VPN such as WireGuard, or a reverse proxy with TLS. Never bind to `0.0.0.0` on a machine reachable from the internet.
- **Prompt injection.** Messages, mail and chat names are written by other people. A message can contain text that tries to make the model call tools. Keep human approval on in your client for every tool that sends anything, and read drafts before you confirm them.
- **WhatsApp uses an unofficial client.** WAHA drives WhatsApp Web. That may violate WhatsApp's Terms of Service and can get the number banned. Use it at your own risk, ideally with a number you can afford to lose.
- **WhatsApp sends only after confirmation.** Sending is two steps: a draft tool, then a send tool the model should call only when you explicitly say so. That rule lives in the tool descriptions, so your client's approval prompt is the real safeguard.
- **WhatsApp attachments are local files.** Attachments are read from the server host's filesystem, only from inside `WHATSAPP_MCP_ATTACHMENTS_DIR`. File attachments stay disabled until you set it. Point it at a dedicated folder, never at your home directory.
- **Office controls the UI.** The Accessibility permission lets the server process read and control every app on the Mac, not only Outlook and Teams. The tools switch folders and chats on screen, and `teams_read` marks chats as read. Prefer a dedicated user account or VM.
- **Your data goes to your model.** Tool output (names, numbers, addresses, message bodies) is sent to whichever LLM your client uses. Pick a provider or a local model you trust with that data.

## Repository layout

```
mcp-suite/
  pyproject.toml        uv workspace root (members: servers/*) and shared dev tools
  uv.lock               one lock for every server
  ruff.toml             shared lint and format rules
  servers/
    whatsapp/           whatsapp-mcp: src/whatsapp_mcp, tests, deploy
    office/             office-mcp: src/office_mcp, tests
  templates/
    server/             skeleton for a new server (not a workspace member)
    scaffold.py         copies the skeleton into servers/<name>
  .github/workflows/    CI
```

This is a [uv workspace](https://docs.astral.sh/uv/concepts/projects/workspaces/): every directory under `servers/` is its own Python package with its own `pyproject.toml`, and they share one lockfile and one `.venv`. Each server stays installable on its own with `uv tool install ./servers/<name>`. That path ignores the workspace lock, so the dependency ranges in each `pyproject.toml` (for example `mcp>=2.3,<3`) are what keep it working.

## Development

```sh
uv sync --all-packages
```

Run the checks for one server from its directory:

```sh
cd servers/whatsapp
uv run ruff check .
uv run ruff format --check .
uv run mypy src tests
uv run pytest
```

Or from the repository root:

```sh
uv run --package whatsapp-mcp pytest servers/whatsapp
```

CI ([.github/workflows/ci.yml](.github/workflows/ci.yml)) discovers every directory in `servers/` and runs ruff, ruff format, mypy (strict) and pytest for each one on `ubuntu-latest`, plus any extra runners listed in an optional `servers/<name>/ci.json` such as `{"runners": ["ubuntu-latest", "macos-latest"]}`. It also scaffolds a throwaway server from the template and checks it, verifies `uv.lock` is current, scans history for secrets with gitleaks, rejects em dashes, and checks that every server ships its required files.

See [CONTRIBUTING.md](CONTRIBUTING.md) for the house rules.

## Adding a new server

1. **Scaffold it.** From the repository root, with `<name>` in lowercase letters and digits and the next free port from the table above:

   ```sh
   uv run --no-project python templates/scaffold.py notes --port 8770
   uv sync --all-packages
   ```

   That creates `servers/notes` with the `notes-mcp` distribution, the `notes_mcp` package, the `notes-mcp` command and the `NOTES_MCP_*` settings, all wired up and passing their tests. `--port` is required, and the script refuses a port another server already uses as its default. `uv sync` adds the new member to `uv.lock`: commit the updated `uv.lock` together with `servers/notes`, or the CI lock check fails.
2. **Implement tools under `features/`.** One directory per domain (`features/inbox/`, `features/drafts/`), each owning its service, models and tool registration. Talk to the outside world through a small client class you can replace with a fake in tests.
3. **Extend settings.** Add server-specific variables to `shared/settings.py`, validated into typed values with a clear `ConfigurationError`, and inject them into services. Never read the environment or the timezone at import time.
4. **Write tests** that need no network and no real account. Use fakes or `httpx.MockTransport`, and only invented names, numbers and addresses.
5. **Document it.** Fill in the server `README.md` (tools table, setup, configuration, security) and `.env.example` (placeholders only).
6. **Register it.** Add a row to the catalog table in this README and claim the port. Add `servers/<name>/ci.json` if it needs more runners than `ubuntu-latest`.
7. **Open a PR** using the checklist in the PR template.

### Server contract

Every server must:

- [ ] Live in `servers/<name>/` with distribution `<name>-mcp`, import package `<name>_mcp` and console script `<name>-mcp`.
- [ ] Use a feature-first layout: `features/<domain>/`, `shared/settings.py`, `shared/errors.py` with custom error classes.
- [ ] Use `uv_build`, depend on `mcp>=2.3,<3`, declare `license = "MIT"` and `requires-python = ">=3.13"`, extend `../../ruff.toml`, and pass `mypy --strict`.
- [ ] Read configuration only from `<NAME>_MCP_*` variables (plus documented upstream ones), with ports and other numbers validated instead of crashing on `int()`.
- [ ] Support both transports: stdio and streamable HTTP at `/mcp` with `stateless_http=True` and `json_response=True`, bound to `127.0.0.1` by default, selectable with `--transport` and `<NAME>_MCP_TRANSPORT`.
- [ ] Support `<NAME>_MCP_TOKEN` (bearer middleware on every HTTP route, constant-time compare) and `<NAME>_MCP_ALLOWED_HOSTS` (passed to the SDK's `TransportSecuritySettings`).
- [ ] Resolve `<NAME>_MCP_TIMEZONE` once in settings and inject it into services.
- [ ] Build instructions and tool descriptions from functions that take the owner label (`<NAME>_MCP_OWNER_NAME`, falling back to "the user").
- [ ] Set accurate `ToolAnnotations` on every tool: `read_only_hint`, `destructive_hint`, `idempotent_hint`, `open_world_hint`.
- [ ] Use draft-then-confirm for anything that sends, writes or deletes, and say so in the tool description.
- [ ] Contain no personal data and no secrets in code, tests, fixtures or docs.
- [ ] Contain no em dashes and no code comments.
- [ ] Use conventional commits with the server name as scope, for example `feat(notes): add search tool`.
- [ ] Pass CI on every runner it declares.

## License

[MIT](LICENSE)
