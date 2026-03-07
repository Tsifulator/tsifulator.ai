# Phase Gate Usage

Before running the gate, complete docs/phase-template.md for the current phase.

Run these at the end of each OpenClaw phase:

```powershell
npm run gate -- --Phase "phase-1-foundation"
npm run gate -- --Phase "phase-2-backend-core"
npm run gate -- --Phase "phase-3-terminal-client"
```

What it does:
- Runs available scripts in this order: lint, typecheck, test, build
- Skips missing scripts without failing
- Fails immediately on first failing check

Tip:
- Keep script names standard (`lint`, `typecheck`, `test`, `build`) so the gate can enforce quality automatically.
- Keep one copy of the completed phase template per phase in docs/phase-reports/.
