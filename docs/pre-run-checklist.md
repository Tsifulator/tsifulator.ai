# tsifulator.ai pre-run checklist (Windows)

Run this once before pasting your big OpenClaw prompt:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\preflight.ps1
```

If preflight is green, run:

```powershell
openclaw --version
Get-Content .\docs\openclaw-master-prompt.md
Get-Content .\docs\openclaw-first-task.md
```

Recommended prompt mode for your use case:
- Keep prompt strict and phase-based.
- Require file-by-file outputs.
- Require exact commands and Windows-safe alternatives.
- Require an acceptance checklist after each phase.

Safety defaults for terminal apply actions:
- Require explicit confirmation for command execution.
- Block high-risk patterns by default (`rm -rf`, destructive wildcard deletes, privilege escalation).
- Keep command output capture bounded and redact obvious secrets.
