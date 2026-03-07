AUTONOMOUS_EXECUTOR MODE.
No clarifying questions. No restating requirements. Start implementation now.

Goal:
Build tsifulator.ai Phase 1 Terminal-First Beta only.

Scope lock:
- In scope: backend + terminal client + sqlite + streaming + apply safety.
- Out of scope: excel/office/ppt/rstudio/billing/sso/enterprise/multi-tenant.

Required endpoints:
- GET /health
- POST /auth/dev-login
- POST /chat
- GET /chat/stream

Safety policy:
- shell_command actions only
- classify command risk: safe|confirm|blocked
- never execute blocked commands
- require explicit confirmation for safe/confirm

Execution format (strict):
SECTION 2: COMMANDS
<exact commands only>

SECTION 3: FILES
FILE: <path>
<full content>
END_FILE

Constraints:
- Max 3 files per response.
- No commentary or prose.
- Continue from latest checkpoint automatically.
- If a validation fails, fix and retry.

Stop condition:
- All Phase 1 acceptance checks pass.
- When all checks pass, print exactly: ACCEPTANCE_STATUS: PASS
