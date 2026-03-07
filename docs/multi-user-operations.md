# Multi-User Operational Considerations

## Current architecture

- **Single SQLite database** — all users share one DB file
- **User isolation** — enforced at query level (all queries filter by `user_id`)
- **Rate limiting** — per-user, configurable via `RATE_LIMIT_MAX_PROMPTS`
- **No shared state** — users cannot see each other's sessions, messages, or proposals

## Scaling thresholds

| Users | Concern | Mitigation |
|---|---|---|
| 1-100 | None | SQLite handles this easily |
| 100-1000 | Write contention | WAL mode (already enabled), consider read replicas |
| 1000-10000 | DB size, query speed | Add indexes, archive old events, consider PostgreSQL migration |
| 10000+ | Single-server limits | Migrate to PostgreSQL + horizontal scaling |

## SQLite multi-user safety

SQLite with WAL mode (already configured) supports:
- Concurrent reads: unlimited
- Concurrent writes: serialized (one at a time, but fast)
- Typical write latency: <1ms for simple inserts

**Risk**: Under heavy concurrent writes (>100 users simultaneously chatting), write queue could back up. Monitor via stream success rate and chat latency metrics.

## Operational runbook

### Adding a new user
No admin action needed — users self-register via `POST /auth/dev-login` or API keys.

### Blocking a user
Not yet implemented. Workaround: delete their API keys and change the dev-login token format.

**Future**: Add `disabled_at` column to users table (defined in `docs/user-model-production.md`).

### Monitoring user activity
- `GET /telemetry/counters` — aggregate KPIs
- `GET /telemetry/recent-events` — recent activity across all users
- Beta scripts: `npm run beta:engagement`, `npm run beta:churn`, etc.

### Database maintenance
- SQLite WAL checkpoint: automatic (SQLite handles this)
- Vacuum: Run `VACUUM` quarterly if DB grows large
- Backup: Copy the `.db` file (safe during reads, pause writes briefly)

### Data isolation audit
All queries that return user-facing data include `WHERE user_id = ?` or `WHERE s.user_id = ?`. The auth middleware (`requireDevAuth`) injects `authUser` into every authenticated request. No endpoint returns cross-user data.

## Future multi-tenancy features

| Feature | Priority | Notes |
|---|---|---|
| User disable/ban | High | Add `disabled_at` check in auth middleware |
| Admin dashboard | Medium | Read-only view of telemetry + user list |
| Team/org support | Low | Shared sessions within a team |
| Data export (GDPR) | Medium | Export all user data as JSON |
| Data deletion (GDPR) | Medium | Cascade delete user + sessions + messages |
