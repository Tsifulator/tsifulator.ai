# Local Development Runbook

## Prerequisites

- **Node.js** v22+ (recommended: use [nvm](https://github.com/nvm-sh/nvm) or [nvm-windows](https://github.com/coreybutler/nvm-windows))
- **pnpm** (install: `npm install -g pnpm`)

## Setup

```sh
# Clone and install
git clone <repo-url> && cd Tsifulator.ai
pnpm install

# Create env file
cp .env.example .env
# Add your OPENAI_API_KEY to .env
```

## Run

```sh
# Development server (auto-reload)
npm run dev

# Production-like start
npm start
```

Server runs at **http://localhost:4000** by default.

## Test

```sh
npm test          # Run all tests (Node.js built-in test runner)
npm run build     # Type-check (tsc --noEmit)
npm run lint      # ESLint
npm run format:check  # Prettier check
```

## Key endpoints

| Endpoint | Method | Auth | Description |
|---|---|---|---|
| `/health` | GET | No | Health check |
| `/auth/dev-login` | POST | No | Dev login (returns token) |
| `/chat` | POST | Yes | Non-streaming chat |
| `/chat/stream` | GET | Yes | SSE streaming chat |
| `/actions/approve` | POST | Yes | Approve/reject command proposals |
| `/sessions` | GET | Yes | List user sessions |
| `/sessions/search` | GET | Yes | Search sessions |
| `/sessions/:id/messages` | GET | Yes | Session message history |
| `/sessions/:id/events` | GET | Yes | Session events |
| `/telemetry/counters` | GET | Yes | KPI counters |
| `/feedback` | POST | Yes | Submit feedback |

## Terminal client

```sh
npm run cli
# Or with KPI mode:
npm run cli -- --kpi [--email user@example.com] [--json]
```

## Project structure

```
server/src/       # Fastify backend
  index.ts        # Routes and server setup
  auth.ts         # Authentication middleware
  chat-engine.ts  # Chat/proposal logic
  config.ts       # Env var validation (zod)
  db.ts           # SQLite (better-sqlite3)
  rate-limit.ts   # Rate limiting
  risk.ts         # Command risk classification + redaction
  types.ts        # Shared types
  adapters/       # Adapter interfaces (terminal, Excel)
clients/terminal/ # Terminal CLI client
tests/            # Integration tests (.test.mjs)
data/             # SQLite database (gitignored)
```

## Environment variables

See `.env.example` for all options. Key ones:

- `OPENAI_API_KEY` — Required for AI features
- `PORT` — Server port (default: 4000)
- `DB_PATH` — SQLite path (default: `./data/tsifulator.db`)
- `JWT_DEV_SECRET` — **Must change in production**
- `RATE_LIMIT_MAX_PROMPTS` — 0 in dev (unlimited), 50 in prod

## Database

SQLite via `better-sqlite3`. Database auto-creates on first run. No manual migrations needed — schema is applied idempotently at startup.

## Troubleshooting

- **Port in use**: Change `PORT` in `.env`
- **Missing OPENAI_API_KEY**: Chat will fail; set it in `.env`
- **Tests fail with EADDRINUSE**: Tests use random ports (4100-4900); if collisions occur, re-run
