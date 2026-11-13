## What and why

## How it was tested

## Checklist

- [ ] Commits follow `type(scope): description`, with the server name as scope
- [ ] `uv run ruff check .`, `uv run ruff format --check .`, `uv run mypy src tests` and `uv run pytest` pass in every server I touched
- [ ] `uv lock --check` passes
- [ ] No secrets, real names, phone numbers, addresses, hostnames or message content in code, tests, fixtures or docs
- [ ] No code comments and no em dashes
- [ ] The server README and `.env.example` match the code

### New or changed server

- [ ] Lives in `servers/<name>/` as `<name>-mcp` / `<name>_mcp`, with a feature-first layout and custom errors in `shared/errors.py`
- [ ] `pyproject.toml` uses `uv_build`, `mcp>=2.3,<3`, `license = "MIT"`, `requires-python = ">=3.13"`, extends `../../ruff.toml`, and passes strict mypy
- [ ] Configuration comes from `<NAME>_MCP_*` variables, with numbers validated by a clear error
- [ ] stdio and streamable HTTP both work, HTTP at `/mcp` with `stateless_http=True` and `json_response=True`, bound to `127.0.0.1` by default
- [ ] `<NAME>_MCP_TOKEN` protects every HTTP route and `<NAME>_MCP_ALLOWED_HOSTS` reaches the SDK transport security settings
- [ ] The timezone is resolved in settings and injected, never read at import time
- [ ] Instructions and descriptions use the owner label, falling back to "the user"
- [ ] Every tool has accurate `ToolAnnotations`
- [ ] Anything that sends, writes or deletes uses draft-then-confirm
- [ ] Tests need no network and no real account
- [ ] The catalog row and port are in the root README, and `ci.json` lists any extra runners
