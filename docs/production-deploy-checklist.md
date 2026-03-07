# Production Deployment Checklist

## Pre-deploy

- [ ] All tests pass (`npm test`)
- [ ] Type-check clean (`npm run build`)
- [ ] Lint clean (`npm run lint`)
- [ ] `.env` values reviewed for production:
  - [ ] `NODE_ENV=production`
  - [ ] `JWT_DEV_SECRET` changed from default
  - [ ] `OPENAI_API_KEY` set
  - [ ] `CORS_ORIGIN` set to actual frontend origin (not `*`)
  - [ ] `RATE_LIMIT_MAX_PROMPTS` set appropriately (default: 50)
  - [ ] `LOG_LEVEL=info`
- [ ] `DB_PATH` points to a persistent volume
- [ ] Docker image builds successfully (`docker build -t tsifulator .`)
- [ ] Docker container starts and `/health` returns 200

## Deploy

- [ ] Push Docker image to registry
- [ ] Deploy to target host (see "Deploy target" below)
- [ ] Verify `/health` and `/api/health` return 200
- [ ] Verify `/auth/dev-login` works
- [ ] Verify `/chat` returns a response
- [ ] Verify `/chat/stream` streams SSE events
- [ ] Verify rate limiting is active (check `X-RateLimit-*` headers)
- [ ] Verify blocked commands return `blocked` status
- [ ] Check logs for clean startup (no warnings/errors)

## Post-deploy

- [ ] Run a manual end-to-end test: login → chat → propose command → approve → verify output
- [ ] Verify SQLite database persists across container restarts
- [ ] Check HEALTHCHECK is passing (`docker inspect --format='{{.State.Health.Status}}'`)
- [ ] Set up log monitoring (stdout/stderr capture)
- [ ] Document the deployed URL and access method

## Rollback plan

1. Keep previous Docker image tagged (e.g., `tsifulator:previous`)
2. If deploy fails: `docker stop tsifulator && docker run tsifulator:previous`
3. SQLite DB is backward-compatible (schema is additive)

## Deploy target options

| Option | Pros | Cons |
|---|---|---|
| Railway / Render | Zero-ops, auto-deploy from git | Cost at scale, less control |
| VPS (Hetzner/DigitalOcean) | Full control, cheap | Manual setup, maintenance |
| Self-hosted (home server) | Free, full control | Uptime depends on hardware |

## Persistent storage notes

- SQLite DB must be on a persistent volume (not ephemeral container storage)
- Mount `/app/data` to a host directory or named volume
- Example: `docker run -v tsifulator-data:/app/data -p 4000:4000 tsifulator`
