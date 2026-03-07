# Tsifulator.ai

AI-powered terminal sidecar — propose, review, and execute commands with built-in safety rails.

## Quick Start

```bash
# Install dependencies (pnpm recommended — npm has known issues on Windows/OneDrive)
pnpm install
pnpm rebuild better-sqlite3 esbuild

# Start the API server
pnpm run dev          # http://localhost:4000

# In another terminal, launch the CLI client
pnpm run cli

# Run tests
pnpm run test         # or: node --test tests/*.test.mjs
```

## Architecture

```
server/src/         Fastify API server (TypeScript)
  index.ts          Routes: health, auth, chat, stream, approve, telemetry, sessions
  config.ts         Env validation (zod) — enforces real secrets in production
  db.ts             SQLite persistence (better-sqlite3, WAL mode)
  chat-engine.ts    Chat + command proposal logic
  risk.ts           Command risk classification + secret redaction
  auth.ts           Dev auth (Bearer token)
  adapters/         Pluggable adapters (terminal, excel)

clients/terminal/   Interactive CLI client
tests/              Integration tests (node:test)
```

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Health check |
| GET | `/api/health` | Health check (aliased) |
| POST | `/auth/dev-login` | Get auth token |
| POST | `/chat` | Send message (non-streaming) |
| GET | `/chat/stream` | SSE streaming chat |
| POST | `/actions/approve` | Approve/reject command proposal |
| POST | `/feedback` | Submit product feedback |
| GET | `/telemetry/counters` | KPI dashboard |
| GET | `/sessions` | List sessions |
| GET | `/sessions/search` | Search sessions |
| GET | `/sessions/:id/messages` | Session message history |
| GET | `/sessions/:id/events` | Session event log |

## CLI Commands

| Command | Description |
|---------|-------------|
| `/help` | Show all commands |
| `/session` | Show current session id |
| `/session-new` | Start a fresh session |
| `/sessions [limit]` | List recent sessions |
| `/sessions-find <text>` | Search sessions by content |
| `/use <id>` | Switch to existing session |
| `/history [id] [limit]` | View message history |
| `/history-export <file>` | Export history (json/jsonl/text/gzip) |
| `/import-history <file\|->` | Load history into context |
| `/kpi [--json]` | Telemetry counters dashboard |
| `/feedback <text>` | Submit product feedback |
| `/clear` | Reset active session |

## Configuration

Copy `.env.example` to `.env` and configure:

| Variable | Default | Description |
|----------|---------|-------------|
| `NODE_ENV` | `development` | `development` / `production` / `test` |
| `PORT` | `4000` | Server port |
| `HOST` | `0.0.0.0` | Bind address |
| `DB_PATH` | `./data/tsifulator.db` | SQLite database path |
| `JWT_DEV_SECRET` | `change_me_in_beta` | Auth secret (**must change in production**) |
| `OPENAI_API_KEY` | — | OpenAI API key |
| `CORS_ORIGIN` | `*` (dev) / `` (prod) | Allowed CORS origin |
| `LOG_LEVEL` | `debug` (dev) / `info` (prod) | Pino log level |

## Deploy

```bash
# Docker
docker build -t tsifulator .
docker run -p 4000:4000 \
  -e JWT_DEV_SECRET=your-real-secret \
  -e OPENAI_API_KEY=sk-... \
  -v tsifulator-data:/app/data \
  tsifulator
```

## Safety

- Commands are classified as `safe`, `confirm`, or `blocked` before execution
- Destructive patterns (`rm -rf`, `Remove-Item -Recurse -Force`, etc.) are blocked
- Secrets in command output are automatically redacted
- Sessions and proposals are isolated per authenticated user
- Duplicate approvals are rejected (409)

## Tests

```bash
node --test tests/*.test.mjs
```

Covers: health endpoints, dev auth, chat (stream + non-stream), action proposal/approval flow, session isolation, blocked commands, telemetry counters, and secret redaction.

## Beta Operations

Beta analytics scripts are in `scripts/` and accessible via `npm run beta:*`. See `docs/beta-onboarding.md` for the full beta ops runbook.
