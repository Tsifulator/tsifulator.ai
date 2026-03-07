Do not ask clarifying questions. Make reasonable defaults. Start coding immediately.

You are my senior product engineer and architect building tsifulator.ai (Unified AI Sidecar).

Mission:
Build Phase 1 only: Terminal-First Beta with a shared backend brain, real-time streaming responses, structured apply actions, and safe command execution.

Scope lock (hard):
- IN SCOPE: Terminal client + backend APIs + SQLite persistence + auth/dev login + streaming + apply safety.
- OUT OF SCOPE: Excel add-in, Office add-in logic, PowerPoint, RStudio, billing, SSO, enterprise compliance, multi-tenant architecture.

Tech constraints:
- Monorepo structure must remain:
  tsifulator.ai/
    server/
    clients/terminal/
    docs/
- Backend: Node.js + TypeScript + Fastify
- Persistence: SQLite (with migration strategy)
- Streaming: SSE endpoint
- Env: .env support with OPENAI_API_KEY
- Action schema (Phase 1): shell_command only
- Terminal client: Node CLI + minimal TUI behavior (streaming text and confirmation prompts)

Required backend endpoints:
- GET /health
- POST /auth/dev-login
- POST /chat            (non-stream fallback)
- GET /chat/stream      (SSE streaming)

Data requirements:
- Persist sessions and message logs in SQLite
- Persist action proposals + user decision (approved/rejected)
- Add created_at timestamps for auditable event ordering

Apply action policy (mandatory):
- Every shell action must be classified as: safe | confirm | blocked
- blocked patterns include at minimum: rm -rf, recursive wildcard deletes, privilege escalation, dangerous chmod/chown operations
- Never execute blocked commands
- confirm commands require explicit user approval
- safe commands still require explicit “Apply” confirmation
- Execution output capture must be bounded (max size) and redacted for obvious secrets patterns

Terminal client requirements:
- Detect current working directory and include it in request context
- Capture last command output (bounded) and include context in request
- Stream assistant output live from SSE
- When action is returned:
  - Print suggested command
  - Show risk level
  - Ask for explicit confirmation
  - Execute only when allowed and confirmed
- Must work on bash and zsh environments (design accordingly)

Architecture requirements:
- Clean folder structure and modular code
- Shared schema/types for message/action payloads where practical
- Centralized error handling
- Input validation for APIs
- Minimal but clear logging

Execution protocol (strict):
1) Output full repo structure first.
2) Output exact terminal commands to initialize project in VS Code.
3) Generate complete backend scaffold file-by-file with paths.
4) Generate complete terminal client scaffold file-by-file with paths.
5) Generate README with exact run instructions.
6) Do not hand-wave; output real code and concrete commands.

Phase process (mandatory):
- Work in phases and stop after each phase summary.
- At end of every phase, fill docs/phase-template.md completely.
- Save each phase report into docs/phase-reports/<phase-name>.md
- Before proceeding to next phase run:
  npm run gate -- --Phase "<phase-name>"
- If gate fails, fix and rerun before continuing.

Acceptance criteria for Phase 1:
- Backend starts without errors.
- /health returns 200 and status payload.
- /chat returns non-stream response.
- /chat/stream streams SSE chunks end-to-end.
- SQLite contains sessions/messages/actions records.
- Terminal client streams response and supports confirmed apply flow.
- Blocked commands are denied.
- README enables a fresh developer to run the system on Windows.

Output contract per phase:
- File tree delta (added/updated/removed)
- Exact commands executed
- Full contents of every created/updated file
- Validation results (lint/typecheck/test/build/health checks)
- Completed acceptance checklist with pass/fail

Business context (for implementation priorities):
- Product focus is reducing workflow fragmentation by embedding AI directly in tools.
- We compete on cross-environment orchestration and structured execution, not on replacing each vertical tool.
- Prioritize reliability, safety, and fast iteration over feature breadth.

Start now:
- Print the full repo structure.
- Then print initialization commands.
- Then begin Phase 1 implementation immediately.
