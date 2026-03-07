# Local Development Runbook

## Prerequisites

- Node.js ≥ 22 (via nvm-windows or nvm)
- pnpm (`npm i -g pnpm`)
- PowerShell 5+ (built into Windows)

## Quick Start

```bash
# 1. Install dependencies
pnpm install

# 2. Copy env (if .env doesn't exist)
#    Defaults work out of the box for dev
cp .env.example .env   # optional — server uses sensible defaults

# 3. Start dev server (auto-reload)
pnpm dev

# 4. In another terminal, run the CLI
pnpm cli
```

## Common Commands

| Action              | Command                            |
| ------------------- | ---------------------------------- |
| Dev server          | `pnpm dev`                         |
| Production start    | `pnpm start`                       |
| Terminal CLI        | `pnpm cli`                         |
| Run tests           | `pnpm test`                        |
| Lint                | `pnpm lint` / `pnpm lint:fix`      |
| Format              | `pnpm format` / `pnpm format:check`|
| Full build check    | `pwsh scripts/build.ps1`           |
| Build + auto-fix    | `pwsh scripts/build.ps1 -Fix`      |
| Build, skip tests   | `pwsh scripts/build.ps1 -SkipTests`|

## Environment Variables

| Variable                  | Default                  | Description                    |
| ------------------------- | ------------------------ | ------------------------------ |
| `PORT`                    | `3000`                   | Server port                    |
| `HOST`                    | `127.0.0.1`              | Bind address                   |
| `DB_PATH`                 | `./data/tsifulator.db`   | SQLite database path           |
| `LOG_LEVEL`               | `info`                   | Fastify log level              |
| `CORS_ORIGIN`             | (empty = disabled)       | CORS allowed origins           |
| `RATE_LIMIT_MAX_PROMPTS`  | `60`                     | Max prompts per window         |
| `RATE_LIMIT_WINDOW_MS`    | `60000`                  | Rate limit window (ms)         |

## Database

SQLite is stored in `data/`. Migrations run automatically on startup.
To reset: delete the `.db`, `-wal`, and `-shm` files in `data/`.

## Troubleshooting

- **Port in use**: Change `PORT` env or kill the process on that port.
- **better-sqlite3 build error**: Run `pnpm rebuild better-sqlite3`.
- **Tests hang**: Each test file starts its own server on a random port. If a test crashes mid-run, orphan processes may linger — kill them with `taskkill /F /IM node.exe` (be careful this kills all Node).
