# Monitoring & Logging Plan

## Logging (built-in)

The server uses **Fastify's Pino logger** with structured JSON output.

| Setting | Development | Production |
|---|---|---|
| `LOG_LEVEL` | `debug` | `info` |
| Format | JSON (Pino) | JSON (Pino) |
| Output | stdout | stdout (captured by Railway) |

### Key log events

All telemetry is written to the `event_log` SQLite table and queryable via:
- `GET /sessions/:id/events` — events for a specific session
- `GET /telemetry/counters` — aggregate KPIs
- `GET /telemetry/recent-events` — latest events across all sessions (for debugging)

### Event types logged

| Event | When |
|---|---|
| `chat_non_stream_completed` | Non-streaming chat response sent |
| `chat_stream_started` | SSE stream opened |
| `chat_stream_completed` | SSE stream finished (includes latency, chunk count, disconnect status) |
| `action_executed` | Command approved and executed |
| `action_rejected` | Command rejected by user |
| `action_blocked` | Blocked command attempted |
| `action_failed` | Command execution error |
| `user_feedback` | User submitted feedback |

## Health checks

| Endpoint | Purpose |
|---|---|
| `GET /health` | Basic liveness check |
| `GET /api/health` | Alias for `/health` |
| Docker `HEALTHCHECK` | Runs every 30s, 3 retries, 10s start period |

Railway automatically monitors the Docker HEALTHCHECK and restarts on failure.

## Monitoring approach (Railway)

### Phase 1 — Launch (built-in, free)

1. **Railway dashboard** — CPU, memory, network metrics
2. **Railway logs** — stdout/stderr captured automatically, searchable
3. **HEALTHCHECK** — auto-restart on failure
4. **`/telemetry/counters`** — check KPIs via CLI: `npm run cli -- --kpi`

### Phase 2 — Post-launch (if needed)

If usage grows or issues arise, add:

1. **Uptime monitoring** — [UptimeRobot](https://uptimerobot.com) (free tier: 50 monitors, 5-min intervals)
   - Monitor `GET /health` — alert on downtime
   - Monitor `GET /api/health` — redundant check

2. **Error alerting** — Add a simple error-count endpoint or pipe Pino logs to a service:
   - [Logtail](https://betterstack.com/logtail) (free tier: 1GB/month)
   - [Axiom](https://axiom.co) (free tier: 500GB ingest/month)

3. **APM** — Only if debugging performance:
   - [Highlight.io](https://highlight.io) (free tier available)

## Key metrics to watch

| Metric | Source | Alert threshold |
|---|---|---|
| Health endpoint status | UptimeRobot | Any non-200 |
| Stream success rate | `/telemetry/counters` | Below 90% |
| Median chat latency | `/telemetry/counters` | Above 5000ms |
| Blocked command attempts | `/telemetry/counters` | Spike (potential abuse) |
| Daily active users | `/telemetry/counters` | Drop to 0 (service down?) |

## Incident response

1. Check Railway logs for errors
2. Hit `/health` — is the server up?
3. Hit `/telemetry/recent-events?limit=20` — what happened recently?
4. Check `/telemetry/counters` — any anomalies in KPIs?
5. If DB corrupt: restore from Railway volume backup or redeploy with fresh DB
