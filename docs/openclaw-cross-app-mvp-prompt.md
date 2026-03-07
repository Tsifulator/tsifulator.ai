AUTONOMOUS_EXECUTOR MODE.
No clarifying questions. No restating requirements. Start implementation now.

Project:
tsifulator.ai

Mission:
Build launchable Phase 1 only: Terminal + Excel-ready orchestration backend with shared memory, safety policy, and auditable action execution.

Scope lock (hard):
IN SCOPE:
- Backend API (Node.js + TypeScript + Fastify)
- SQLite persistence with migration path
- Terminal adapter end-to-end
- Excel adapter interface scaffold (no deep Office UI implementation)
- Shared session/message/action/audit model
- SSE streaming + non-stream chat
- Apply safety policy with risk classes

OUT OF SCOPE:
- Custom model training
- Billing, SSO, enterprise compliance
- Full Office add-in UX
- PowerPoint and RStudio implementation in this phase

Required endpoints:
- GET /health
- POST /auth/dev-login
- POST /chat
- GET /chat/stream
- POST /actions/approve
- GET /sessions/:id/events

Data model (required tables/entities):
- users
- sessions
- messages
- action_proposals
- action_executions
- approvals
- event_log
- adapter_states
All entities include created_at.

Safety policy (mandatory):
- Action type: shell_command only in Phase 1
- Risk classification: safe | confirm | blocked
- blocked examples include at minimum:
  - rm -rf
  - wildcard recursive deletes
  - privilege escalation (sudo/runas/elevation)
  - dangerous chmod/chown patterns
- Never execute blocked actions
- safe and confirm both require explicit approval
- Execution output must be bounded and redacted for obvious secrets

Architecture constraints:
- Clean modular folder structure under server/
- adapter contract must exist:
  - captureContext()
  - proposeActions()
  - validateAction()
  - executeAction()
  - emitEvents()
- terminal adapter fully functional in Phase 1
- excel adapter skeleton with testable mocked flow
- centralized error handling
- request validation on all write APIs

Execution protocol (strict):
1) Inspect existing repository and continue from latest checkpoint.
2) Implement in small increments.
3) Max 3 files per response.
4) After each increment, run relevant validation and fix failures.
5) Keep outputs deterministic and concise.

Output format (strict):
SECTION 2: COMMANDS
<exact commands only>

SECTION 3: FILES
FILE: <path>
<full content>
END_FILE

No prose outside the required sections.

Completion criteria:
- Backend runs successfully
- Required endpoints respond correctly
- SQLite persistence works for sessions/messages/actions/events
- Terminal adapter can stream, propose actions, and enforce approvals
- README run steps are accurate on Windows
- When all criteria pass print exactly:
ACCEPTANCE_STATUS: PASS