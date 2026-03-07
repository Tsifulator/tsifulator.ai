# Analytics & Events Plan

## Current state

All user activity is already captured in the `event_log` SQLite table via `db.logEvent()`. This is the foundation for analytics.

## Events already tracked

| Event | Data captured |
|---|---|
| `chat_non_stream_completed` | hasProposal, latencyMs |
| `chat_stream_started` | hasProposal |
| `chat_stream_completed` | chunks, hasProposal, firstTokenLatencyMs, totalLatencyMs, clientClosedEarly |
| `action_executed` | proposalId |
| `action_rejected` | proposalId |
| `action_blocked` | proposalId |
| `action_failed` | proposalId, output |
| `user_feedback` | userId, text, source |
| `terminal_context_captured` | cwd, hasLastOutput |
| `excel_adapter_scaffold_used` | sessionId |
| `rstudio_adapter_scaffold_used` | sessionId |

## Aggregated metrics (via `/telemetry/counters`)

- New beta users (7d)
- Daily active users
- Prompts sent (all-time + 24h)
- Apply actions proposed/confirmed
- Blocked command attempts
- Stream requests/completions/success rate
- Median chat latency
- Median stream first-token latency

## Phase 1 — Launch analytics (built-in)

No external services needed. Use existing infrastructure:

1. **KPI dashboard**: `npm run cli -- --kpi` or `GET /telemetry/counters`
2. **Recent events**: `GET /telemetry/recent-events?limit=50`
3. **Per-session drill-down**: `GET /sessions/:id/events`
4. **Beta scripts**: Existing `npm run beta:*` scripts for engagement, churn, retention, sentiment

## Phase 2 — Product analytics (post-launch)

When user base grows, add lightweight client-side analytics:

### Option A: PostHog (recommended)
- Open-source, self-hostable
- Free tier: 1M events/month
- Feature flags, funnels, retention charts
- Integration: Add `posthog-node` server-side, emit events alongside `db.logEvent()`

### Option B: Plausible + custom
- Plausible for web dashboard traffic (if web UI added)
- Keep server-side events in SQLite, build custom dashboards

## Events to add (post-launch)

| Event | When | Purpose |
|---|---|---|
| `user_registered` | New user created | Track growth |
| `user_upgraded` | Plan change | Track conversion |
| `api_key_created` | API key issued | Track CLI adoption |
| `api_key_revoked` | API key revoked | Track churn signal |
| `rate_limit_hit` | 429 returned | Track upgrade triggers |
| `session_resumed` | Existing sessionId reused | Track engagement depth |

## Key funnels to track

1. **Activation**: Register → First prompt → First command approved → Return next day
2. **Upgrade**: Free user → Hit rate limit → View upgrade CTA → Convert to Pro
3. **Retention**: D1 / D7 / D30 return rates

## Data retention

- Event log: Keep indefinitely in SQLite (small footprint)
- If scaling: Archive events older than 90 days to compressed JSON files
- GDPR: Add user data export/deletion endpoint when needed
