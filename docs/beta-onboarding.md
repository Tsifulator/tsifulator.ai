# Tsifulator.ai Beta Onboarding (Windows)

Use this runbook to onboard a new beta tester in under 10 minutes.

## 1) Run one-command onboarding checks

```powershell
npm run beta:onboard
```

Optional flags:

```powershell
npm run beta:onboard -- -Email tester@company.com
npm run beta:onboard -- -Owner ntsif
npm run beta:onboard -- -SkipInstall
npm run beta:onboard -- -StartServer
npm run beta:onboard -- -RunTrafficWarmup
npm run beta:onboard:warmup
```

What this validates:
- Node and npm availability
- `.env` presence (auto-creates from `.env.example` when possible)
- dependencies install
- typecheck (`npx tsc --noEmit`)
- tests (`npm test`)
- optional live traffic warmup (`/chat` + `/chat/stream`) with latency-aware KPI check
- automatic checkpoint append to `docs/daily-triage/<today>.md` after warmup

## 2) Start app for beta user

In terminal A:

```powershell
npm run dev
```

In terminal B:

```powershell
npm run cli
```

## 3) First beta session checklist

In CLI:
- Login with tester email
- Send one prompt (for example: `list files`)
- Run `/kpi` to verify counters endpoint responds
- Run `/feedback first beta session complete`

## 4) Daily operating loop

- Generate today's triage file:

```powershell
npm run beta:triage:new
```

- Start each day with `/kpi`
- Capture user sentiment with `/feedback <text>`
- Export history when needed with `/history-export ...`

## 5) Troubleshooting quick hits

- If API fails to start, run:

```powershell
npx tsc --noEmit
npm test
```

- If autonomous runs look stuck, check:

```powershell
npm run openclaw:trust-gate
npm run openclaw:overnight:status
```
