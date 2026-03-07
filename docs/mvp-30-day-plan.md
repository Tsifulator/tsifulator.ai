# Tsifulator.ai 30-Day MVP Plan (Terminal-First)

## Success target by Day 30
- Usable Terminal-First beta with streaming chat, safe Apply actions, SQLite logs, and 5-10 real users testing.

## North-star KPI
- Weekly active beta users executing at least 1 confirmed Apply action per day.

## Week 1 (Days 1-7): Foundation + Core Backend
Goals:
- Stable backend skeleton and database foundation.
- Streaming and non-stream chat endpoints live.

Deliverables:
- Fastify TypeScript server with modular architecture.
- Endpoints: /health, /auth/dev-login, /chat, /chat/stream.
- SQLite schema for users/sessions/messages/actions.
- .env loading + config validation.
- Basic logging and centralized error handling.

Acceptance checks:
- Local run works from clean clone.
- SSE stream emits chunks correctly.
- Message/session/action records persist.

## Week 2 (Days 8-14): Terminal Client + Safety Layer
Goals:
- CLI client streams responses and can Apply safely.

Deliverables:
- Node CLI app with minimal TUI behavior.
- CWD detection and bounded last-output context capture.
- Action renderer with risk labels: safe/confirm/blocked.
- Confirmation flow before execution.
- Blocklist for dangerous commands.
- Output redaction + max output limits.

Acceptance checks:
- End-to-end prompt -> stream -> apply flow works.
- Blocked commands never execute.
- Confirm flow required for all command execution.

## Week 3 (Days 15-21): Reliability + UX Polish + Observability
Goals:
- Make beta reliable enough for external users.

Deliverables:
- Retry/error states in terminal client.
- Better session continuity and message history retrieval.
- Structured logs and basic metrics counters.
- Smoke tests for core routes and action policy.
- Developer docs and troubleshooting guide.

Acceptance checks:
- 1-hour manual soak test with no critical crashes.
- Core happy-path works repeatedly across sessions.

## Week 4 (Days 22-30): Beta Launch + Feedback Loop
Goals:
- Ship to first users, learn fast, improve quickly.

Deliverables:
- Private beta onboarding doc + script.
- 5-10 design partners onboarded.
- In-app feedback capture command (e.g., /feedback).
- Daily issue triage and patch releases.
- Prioritized backlog for Phase 2 (Excel bridge prep only, no implementation yet).

Acceptance checks:
- At least 5 active beta users in one week.
- At least 20 confirmed Apply actions executed safely.
- Top 10 UX issues documented and prioritized.

## Day-by-day practical sprint checklist
Days 1-2:
- Project scaffolding, env/config, health route, auth dev login.

Days 3-4:
- /chat and /chat/stream with OpenAI integration.

Days 5-7:
- SQLite models + migrations + persistence wiring.

Days 8-10:
- Terminal CLI skeleton + streaming renderer.

Days 11-12:
- Apply pipeline + risk classifier + confirmation UI.

Days 13-14:
- Blocklist hardening + bounded output + redaction.

Days 15-17:
- Reliability fixes, reconnection behavior, better errors.

Days 18-19:
- Session history handling and UX polish.

Days 20-21:
- Smoke tests + docs + runbooks.

Days 22-24:
- Beta onboarding, telemetry counters, feedback channel.

Days 25-27:
- Live user support and high-priority fixes.

Days 28-30:
- Retrospective, KPI review, Phase 2 scope lock.

## Minimal KPI dashboard to track daily
- New beta users onboarded
- Daily active users
- Prompts sent
- Streaming success rate
- Apply actions proposed
- Apply actions confirmed
- Blocked command attempts
- Error rate per endpoint
- Median latency (/chat and /chat/stream first token)

## Founder operating rhythm (daily)
- 30 min: KPI + error review
- 90 min: top bug/UX fix
- 60 min: user interviews/support
- 60 min: roadmap and backlog grooming
- End of day: release note + next-day priority lock

## Day 2 progress snapshot (2026-03-02)
Status:
- On track for Week 1/2 engineering scope, with selected Week 3/4 operations work already in place.

Delivered today:
- In-app feedback loop implemented (`POST /feedback` + CLI `/feedback`).
- KPI counters implemented (`GET /telemetry/counters` + CLI `/kpi`).
- Beta onboarding assets implemented (`npm run beta:onboard`, `docs/beta-onboarding.md`).
- Daily triage system implemented (`npm run beta:triage:new`, `docs/daily-triage/2026-03-02.md`).
- Safety validation executed (blocked command simulation confirmed policy enforcement).

Latest KPI baseline (from daily run):
- New beta users (7d): 4
- Daily active users: 3
- Prompts sent (all-time / 24h): 3 / 3
- Apply actions proposed / confirmed: 3 / 2
- Blocked command attempts: 1
- Stream requests / completions: 1 / 1
- Stream success rate: 100%

Known risk:
- External beta-user volume is not started yet; KPI quality still reflects internal validation traffic.

Reliability milestone:
- `npm run openclaw:trust-gate` now passes with `READY_FOR_TIER2=true`.
- DB migration self-heal for legacy duplicate approval/execution rows is implemented to prevent startup lockups.

Next 24h priorities:
- Run at least one external beta-user session and capture 3+ feedback entries.
- Log first day-over-day KPI checkpoint delta via `npm run beta:checkpoint:append`.
- Prepare design-partner outreach/onboarding slots for Week 4 launch target.
