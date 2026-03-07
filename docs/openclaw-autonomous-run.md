# OpenClaw Autonomous Run (Windows)

This mode runs OpenClaw from terminal in repeated turns with retry/backoff.

## 1) Start gateway in one terminal

```powershell
openclaw gateway run --force
```

Keep this terminal open.

## 2) In another terminal, run project preflight

```powershell
cd "C:\Users\ntsif\OneDrive\Tsifulator.ai"
npm run preflight
```

## 3) Start autopilot loop

```powershell
npm run openclaw:auto
```

## 3b) Preferred for hours-long runs (auto-restart supervisor)

```powershell
npm run openclaw:supervisor
```

Quick smoke mode:

```powershell
npm run openclaw:supervisor:quick
```

## 3c) One-click overnight mode (recommended before sleep)

```powershell
npm run openclaw:overnight:start
```

Check progress:

```powershell
npm run openclaw:overnight:status
```

Stop overnight mode:

```powershell
npm run openclaw:overnight:stop
```

## 3d) Tier-2 trust gate (before long autonomous runs)

```powershell
npm run openclaw:trust-gate
```

Interpretation:
- `READY_FOR_TIER2=true` means environment + gateway + typecheck/tests + dry-run + quick supervisor check all passed.
- `READY_FOR_TIER2=false` means do not rely on unattended long runs yet; fix failing checks first.

Optional custom run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\openclaw-autopilot.ps1 -PromptFile .\docs\openclaw-autopilot-prompt.md -SessionId tsif-phase1 -MaxTurns 180 -TurnDelaySeconds 25 -RateLimitCooldownSeconds 180
```

## 4) Monitor outputs

- Runtime logs are written under `docs/phase-reports/autopilot-runs/`.
- If you hit rate limits, the script cools down and retries.
- Supervisor mode re-runs autopilot cycles automatically and re-checks gateway health before each cycle.

Overnight tips:
- Keep the laptop plugged in.
- Set Windows power mode to prevent sleep while plugged in.
- Do not close the laptop lid unless lid-close action is set to "Do nothing".

## Notes

- Funds and rate limits are different. You can have balance and still hit provider RPM/TPM limits.
- Keep one active coding session at a time to reduce throttling.
- Prefer small output batches (3 files max per turn) for stability.
