# Tsifulator.ai Cross-App 60-Day Launch Plan

## Strategic Positioning
- Do not compete on raw model quality.
- Compete on cross-app execution, shared context, and safe actions.
- Use best available app copilots/engines where they exist.
- Build first-party capability only where there is a workflow gap (initially RStudio workflows).

## Product Thesis
Tsifulator.ai is a unified orchestration layer across tools, not a replacement for each tool-native copilot. The moat is:
- One shared memory across apps
- One safety and approval policy layer
- One action/audit log and governance surface
- One workflow graph that spans tools

## 60-Day Outcome Targets
By Day 60:
- 1 production-ready backend brain
- 2 active adapters in real usage (Terminal + Excel)
- 1 early RStudio companion MVP (guided workflows)
- 10 design partners onboarded
- 1000+ cross-app events captured in unified DB
- 80%+ successful action completion rate on core workflows

## Scope Assumptions (for this plan)
- Team is lean and speed-focused
- Initial user segment: analysts, ops, and technical founders
- Launch starts as private beta
- Compliance scope is basic startup-grade controls, not enterprise certifications

## Architecture Choices (Recommended)

### 1) Core Backend
- Stack: Node.js + TypeScript + Fastify
- Data: SQLite for beta, migration path to Postgres
- API style: REST + SSE streaming
- Modules:
  - auth
  - sessions
  - messages
  - actions
  - adapter registry
  - policy engine
  - telemetry

### 2) Unified Data Model
- users
- sessions
- messages
- action_proposals
- action_executions
- approvals
- app_context_snapshots
- event_log
- adapter_states

Principles:
- Every action is traceable
- Every assistant decision is auditable
- Cross-app context is queryable by session and user

### 3) Adapter Interface (Core Moat)
Each app integration implements the same contract:
- captureContext(): current app state and bounded history
- proposeActions(): app-specific executable actions
- validateAction(): classify as safe, confirm, blocked
- executeAction(): run only after user approval
- emitEvents(): standardized lifecycle events to backend

Initial adapters:
- terminal adapter (high leverage, fastest to ship)
- excel adapter (copilot-aware integration path)
- rstudio companion adapter (first-party guided actions)

### 4) Safety and Policy Engine
- command risk classification: safe, confirm, blocked
- hard block patterns for destructive actions
- required user confirmation for all execution
- output size limits and secret redaction
- full action audit trail

### 5) Model Strategy
- provider-agnostic model gateway
- route by cost and latency profile
- fallback path on provider errors/rate limits
- no model training in first 60 days

## 60-Day Milestone Plan

### Days 1-10: Foundation and Control Plane
Goals:
- Make unified backend production-usable for private beta.

Deliverables:
- Fastify backend scaffold with modular architecture
- Auth and session APIs
- Message and action persistence
- SSE chat stream endpoint
- Policy engine v1 (safe/confirm/blocked)
- Adapter registry and terminal adapter baseline

Exit Criteria:
- End-to-end terminal request -> stream -> action proposal -> approval -> execution -> audit record works reliably.

### Days 11-20: Terminal Experience and Reliability
Goals:
- Make terminal flow sticky and trustworthy.

Deliverables:
- Better CLI UX (streaming, confirmations, retries)
- Bounded context capture (cwd + last output)
- Execution trace viewer in logs/report endpoint
- Rate-limit resilience and queue controls
- Health and smoke tests

Exit Criteria:
- 1-hour soak test passes with no critical failures.

### Days 21-35: Excel Integration Layer
Goals:
- Add practical Excel-side value without replacing native Excel copilot.

Deliverables:
- Excel adapter proof-of-value (context in, actions out)
- Shared session linking between terminal and Excel workflows
- Cross-app memory retrieval API
- Action normalization so Excel and terminal actions share audit model

Exit Criteria:
- At least 3 real tasks completed across Excel + terminal in single session flows.

### Days 36-50: RStudio Companion MVP
Goals:
- Fill the gap where there is weaker default copiloting.

Deliverables:
- RStudio companion action model (script/run/check workflows)
- Context capture and bounded output integration
- Safety policy inheritance from core engine
- Workflow templates for common analysis tasks

Exit Criteria:
- At least 2 repeatable RStudio workflows complete with approvals and audit logs.

### Days 51-60: Beta Launch and Expansion Readiness
Goals:
- Launch to design partners and validate retention signals.

Deliverables:
- Private beta onboarding package
- Usage analytics dashboard (core KPIs)
- Issue triage and patch cadence
- Adapter SDK draft for next app integrations
- Expansion backlog ranked by ROI and integration feasibility

Exit Criteria:
- 10 design partners onboarded
- Weekly active usage with repeat sessions
- Clear top-3 expansion app candidates selected

## Pricing Hypothesis (Initial)

### Packaging
- Free pilot: limited sessions per week, no team sharing
- Pro individual: unlimited personal sessions, advanced history and action logs
- Team beta: shared workspace memory, role-based approvals, team audit exports

### Pricing Test Bands
- Pro: 19 to 39 per user per month
- Team: 99 to 299 per workspace per month (with seat cap)

### What to measure
- Time-to-first-successful-action
- Actions per active user per week
- Cross-app session rate
- 4-week retention of active users
- Conversion from free pilot to paid

## Go-To-Market (First 60 Days)
- Channel: founder-led outbound to analysts/ops power users
- Motion: design partner program with weekly feedback loops
- Positioning message: one copilot memory and action layer across your real tool stack
- Weekly cadence:
  - Monday: backlog and KPI review
  - Tuesday-Thursday: ship and validate
  - Friday: partner interviews and roadmap reprioritization

## KPI Dashboard (Must-Have)
- daily active users
- weekly active users
- successful action executions
- blocked action attempts
- approval-to-execution conversion rate
- stream error rate
- p50 and p95 first-token latency
- cross-app session count
- retention week 1 to week 4

## Risks and Mitigations
- Platform dependency risk:
  - Mitigation: provider-agnostic gateway and adapter abstraction
- Rate-limit instability:
  - Mitigation: queueing, backoff, cached context summaries
- Over-expansion too early:
  - Mitigation: strict adapter admission criteria by ROI
- Trust risk from unsafe actions:
  - Mitigation: confirmation gates, hard blocks, full auditability

## Adapter Expansion Framework (After Day 60)
Score each candidate app from 1 to 5 on:
- User demand
- Integration complexity
- Actionability (can it execute useful actions?)
- Moat contribution (does it improve cross-app value?)
- Revenue impact

Prioritize apps with highest weighted total.

## Immediate Next 7 Days (Concrete)
1) Freeze architecture decisions and schema for unified action/event model.
2) Implement terminal adapter end-to-end with safety gates.
3) Add observability baseline and error budget alerts.
4) Define Excel adapter MVP boundary and first 3 workflows.
5) Recruit first 5 design partners and schedule weekly feedback sessions.

## Decision Rule for Building Your Own Models Later
Only start own model/engine work if all are true:
- clear performance gap blocks retention
- third-party model costs materially hurt margins
- you have enough proprietary workflow data to fine-tune effectively
- engineering focus can shift without slowing core product velocity

Until then, stay focused on orchestration, memory, safety, and cross-app execution quality.