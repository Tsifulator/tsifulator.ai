# Design Partner Session Playbook

Use this for each real beta user session.

## 1) Bootstrap session artifacts

```powershell
npm run beta:partner:start -- -PartnerName "Partner A" -PartnerEmail "partner.a@company.com" -SessionGoal "Validate daily terminal workflow" -Owner "ntsif"
```

Dry-run (no API calls):

```powershell
npm run beta:partner:start -- -PartnerName "Partner A" -DryRun
```

What this does:
- Ensures today's triage file exists
- Appends latest KPI checkpoint into triage
- Creates `docs/partner-sessions/<date>-<partner>.md`
- Submits kickoff `/feedback` event with partner + goal context

## 2) Run the live session

In terminal A:

```powershell
npm run dev
```

In terminal B:

```powershell
npm run cli
```

During session:
- Collect at least 3 user quotes via `/feedback ...`
- Confirm one safe apply action execution
- Confirm one blocked command is denied

## 3) Closeout

- Run checkpoint append:

```powershell
npm run beta:checkpoint:append
```

- Fill:
  - `docs/partner-sessions/<date>-<partner>.md`
  - `docs/daily-triage/<date>.md`

## Success criteria per session
- 1 real workflow completed
- 3+ concrete feedback entries
- At least 1 actionable UX improvement identified
