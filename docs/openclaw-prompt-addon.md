# OpenClaw Prompt Add-on (paste below your main prompt)

Execution policy (strict):
1) Work in phases and stop after each phase summary.
2) For every phase, output:
   - the completed template from docs/phase-template.md
   - full file tree delta
   - exact commands run
   - all file contents (path-by-path)
   - acceptance checklist status (pass/fail)
3) Before moving to next phase, run quality gate:
   - npm run gate -- --Phase "<phase-name>"
4) If gate fails, fix issues before continuing.
5) Never skip safety constraints for terminal command execution.
6) Keep code MVP-only; no extras outside requirements.
7) Do not proceed to the next phase until the phase template is fully completed.

Terminal safety policy:
- Always ask confirmation before executing generated commands.
- Block by default: rm -rf, recursive deletes with wildcards, chmod -R 777, privilege escalation, shell redirection to system paths.
- Redact potential secrets from logs and terminal output previews.

Definition of done for Beta Phase 1:
- Server starts and /health is reachable.
- /chat and /chat/stream both work.
- SQLite persists sessions + messages.
- Terminal client streams responses and supports confirmed apply action.
- README has clean setup + run on Windows.
