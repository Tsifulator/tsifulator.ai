# Production Deployment Runbook

## Overview

tsifulator.ai runs as a single Fastify + SQLite process. Deployment target is flexible — Docker, bare VM, or managed Node hosting.

## Pre-deploy Checklist

- [ ] All tests pass: `pnpm test`
- [ ] Type-check clean: `npx tsc --noEmit`
- [ ] Lint clean: `pnpm lint`
- [ ] Environment variables set (see below)
- [ ] `data/` directory exists and is writable
- [ ] Reverse proxy configured (nginx/caddy) with HTTPS

## Environment Variables (Production)

| Variable                  | Recommended Value               | Notes                        |
| ------------------------- | ------------------------------- | ---------------------------- |
| `PORT`                    | `3000`                          | Behind reverse proxy         |
| `HOST`                    | `0.0.0.0`                       | Accept external connections  |
| `DB_PATH`                 | `/app/data/tsifulator.db`       | Persistent volume mount      |
| `LOG_LEVEL`               | `warn`                          | Reduce noise in prod         |
| `CORS_ORIGIN`             | `https://tsifulator.ai`         | Lock down to your domain     |
| `NODE_ENV`                | `production`                    |                              |
| `RATE_LIMIT_MAX_PROMPTS`  | `30`                            | Tighter in prod              |
| `RATE_LIMIT_WINDOW_MS`    | `60000`                         |                              |

## Docker

```bash
docker build -t tsifulator .
docker run -d \
  --name tsifulator \
  -p 3000:3000 \
  -v tsifulator-data:/app/data \
  -e NODE_ENV=production \
  -e HOST=0.0.0.0 \
  tsifulator
```

Validate the Dockerfile runs end-to-end before first deploy (Phase 3 item).

## Bare VM / VPS

```bash
# Install deps
pnpm install --frozen-lockfile --prod

# Start with process manager
npx pm2 start "node ./node_modules/tsx/dist/cli.mjs server/src/index.ts" \
  --name tsifulator \
  --max-memory-restart 512M

# Or use systemd (create a unit file)
```

## Reverse Proxy (nginx example)

```nginx
server {
    listen 443 ssl;
    server_name tsifulator.ai;

    location / {
        proxy_pass http://127.0.0.1:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_buffering off;           # required for SSE
        proxy_read_timeout 300s;       # keep SSE alive
    }
}
```

## Health Check

```bash
curl http://localhost:3000/health
# Expected: {"status":"ok","service":"tsifulator.ai",...}
```

## Rollback

1. Stop the process (`pm2 stop tsifulator` or `docker stop tsifulator`)
2. Deploy previous version
3. Start again
4. Database is forward-compatible — no rollback migrations needed (yet)

## Monitoring

- Health endpoint: `GET /health` and `GET /api/health`
- Telemetry: `GET /telemetry/counters` (requires auth)
- Logs: stdout via Fastify logger; pipe to your log aggregator
- SQLite WAL size: monitor `data/*.db-wal` — checkpoint runs on shutdown

## Backup

SQLite backup while running:
```bash
sqlite3 /app/data/tsifulator.db ".backup /backups/tsifulator-$(date +%F).db"
```
