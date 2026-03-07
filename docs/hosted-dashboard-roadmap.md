# Hosted Dashboard Roadmap

## Decision: Defer until post-launch

A hosted web dashboard is **not required for launch**. The terminal CLI + API provide full functionality. A dashboard becomes valuable when:

1. Non-technical users need access (team plan)
2. Admin monitoring becomes frequent
3. Onboarding needs a visual entry point

## If/when to build

### Phase 1 — Admin panel (internal)

**When**: After reaching ~100 users
**What**: Read-only dashboard for the founder to monitor:
- Active users, prompts/day, stream success rate
- Recent events feed
- User list with activity timestamps

**Tech**: Simple static HTML page served from `server/public/admin.html`, fetching from existing API endpoints. No framework needed.

### Phase 2 — User dashboard (Pro feature)

**When**: After Pro tier launches
**What**: Each user sees:
- Their usage (prompts today/month)
- Session history browser
- API key management
- Billing/plan management (Stripe portal link)

**Tech**: Lightweight SPA (Preact or vanilla JS) served from `server/public/`. Calls existing authenticated API endpoints.

### Phase 3 — Team dashboard

**When**: If team plan gains traction
**What**: Team admin sees:
- Team member activity
- Shared session access
- Usage allocation per member

**Tech**: Extend user dashboard with team-scoped API endpoints.

## Build vs. buy

| Option | Pros | Cons |
|---|---|---|
| Custom (recommended) | Uses existing API, full control, no extra cost | Dev time |
| Retool/Appsmith | Fast to build, good for admin panels | Monthly cost, external dependency |
| Grafana | Great for metrics dashboards | Overkill for user-facing UI |

## Recommendation

1. **Launch without a dashboard** — CLI + API is sufficient
2. **Post-launch**: Build a minimal admin page (`server/public/admin.html`) using existing endpoints
3. **Only build user dashboard** when Pro plan conversion data shows demand
