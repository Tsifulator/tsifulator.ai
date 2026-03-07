# tsifulator.ai launch checklist

## Launch target
- [x] Public launch-ready in 2 months
- [x] Core product stable
- [x] Terminal sidecar usable end-to-end
- [x] Streaming chat stable
- [x] Approve/confirm execution flow stable
- [x] Persistent sessions working

## Reliability
- [x] Server starts cleanly every time
- [x] /health works
- [x] /api/health works
- [x] SSE streaming is stable
- [x] Env vars are validated
- [x] SQLite startup is safe and repeatable

## Core product
- [x] Main user workflow works end-to-end
- [x] Chat non-stream works
- [x] Chat stream works
- [x] Action proposal flow works
- [x] Action approval flow works
- [x] Telemetry counters work
- [x] Session history works

## Auth and security
- [x] Dev login works
- [x] User/session isolation works
- [x] Dangerous commands are blocked or require confirm
- [x] Secrets are redacted from output

## Tests
- [x] Health endpoint test
- [x] Dev login test
- [x] Chat non-stream test
- [x] Chat stream SSE contract test
- [x] One end-to-end happy path test

## Developer experience
- [x] dev/build/test scripts stable
- [x] lint/prettier configured
- [x] README/runbook updated

## Launch prep
- [x] Production config path exists
- [x] Logging/telemetry exists
- [x] Billing/auth plan is defined
- [x] First deploy path is defined
