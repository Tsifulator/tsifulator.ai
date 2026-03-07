# tsifulator.ai product roadmap

## Mission

Ship tsifulator.ai publicly in 2 months as a Unified AI Sidecar:

- terminal-first AI copilot
- safe apply/confirm execution
- streaming chat
- persistent sessions
- production-ready backend

## Rules

- Prioritize shipping
- Keep changes incremental and testable
- Prefer working functionality over perfect architecture
- Do not rewrite stable code unless necessary
- Complete highest-impact unfinished items first

---

## Phase 1 — Launch-critical product work

### Core workflows
- [x] Terminal client can send a prompt and receive a non-stream chat response
- [x] Terminal client can receive SSE streaming output cleanly
- [x] Terminal client can display action proposals clearly
- [x] Terminal client can approve/reject proposals from terminal
- [x] Approved commands execute and return output safely
- [x] Rejected commands are logged correctly
- [x] Session history persists across runs
- [x] Session search works from terminal workflow

### Safety / execution
- [x] Dangerous shell commands are blocked
- [x] Medium-risk commands require explicit confirm
- [x] Safe commands run smoothly
- [x] Command output is bounded and redacted
- [x] Secrets never appear in terminal output
- [x] Execution timeout behavior is user-friendly

### Streaming / chat
- [x] SSE stream handles disconnects cleanly
- [x] SSE stream sends keepalive heartbeats if needed
- [x] Chat responses are stable under longer prompts
- [x] SessionId behavior is consistent in stream and non-stream modes
- [x] Telemetry events are emitted for stream start/end/failure

---

## Phase 2 — Production readiness

### Reliability
- [x] App starts cleanly with missing/invalid env var protection
- [x] SQLite DB initializes safely every time
- [x] Migrations are idempotent
- [x] Startup logs are readable and actionable
- [x] Server survives invalid requests without crashing
- [x] Health endpoints reflect real service readiness

### Testing
- [x] Expand endpoint coverage beyond 5 existing tests
- [x] Add tests for approve/reject flow
- [x] Add tests for blocked dangerous commands
- [x] Add tests for session persistence
- [x] Add tests for telemetry counters
- [x] Add tests for SSE disconnect behavior
- [x] Add one realistic end-to-end happy path test

### DX
- [x] Add lint config
- [x] Add prettier config
- [x] Add reliable build script
- [x] Add runbook for local development
- [x] Add runbook for production deployment
- [x] Add seed/dev bootstrap helper if useful

---

## Phase 3 — Launch features

### Auth / users
- [x] Replace or extend dev-login toward real auth plan
- [x] Define user model for production
- [x] Ensure strict user/session isolation
- [x] Add auth/session documentation

### Product polish
- [x] Improve terminal UX formatting
- [x] Add clearer proposal display
- [x] Add better error messages for auth failures
- [x] Add better error messages for command failures
- [x] Improve telemetry visibility for debugging

### Deployment
- [ ] Validate Dockerfile end-to-end
- [x] Create production deployment checklist
- [x] Define deploy target
- [x] Add basic monitoring/logging plan
- [x] Run load/stress testing on critical flows

---

## Phase 4 — Post-launch expansion

### Excel / RStudio adapters
- [x] Define adapter interfaces clearly
- [x] Implement Excel adapter skeleton
- [x] Implement RStudio adapter skeleton
- [x] Route shared session logic through adapters
- [x] Add adapter-specific tests

### SaaS expansion
- [x] Add billing strategy document
- [x] Add analytics/events plan
- [x] Add multi-user operational considerations
- [x] Add roadmap for hosted dashboard if needed

---

## Current execution rule

In each iteration:

1. pick the highest-impact unfinished item
2. make one safe, meaningful step toward finishing it
3. run tests/build/dev checks as needed
4. summarize progress
5. move to the next unfinished item
