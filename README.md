# Tsifulator.ai

**Agentic Sandbox for Financial Analysts.** A cross-app AI operating layer
that works inside Excel, RStudio, Word, PowerPoint, Gmail, Notes, and
your browser — with persistent per-project memory and a desktop agent
for Ribbon dialogs Office.js can't reach.

> *"Bloomberg built a terminal for data. We're building a terminal for everything
> an analyst actually does."*

---

## What is this, concretely

- An **Excel add-in** (task pane) that reads your live workbook and executes
  structured actions against it via Office.js.
- A **FastAPI backend** (deployed on Railway) that talks to Claude, emits
  structured actions, and gates them through a phantom-sheet guard,
  formula-literacy rules, and a lock-enforcement layer.
- A **desktop agent** (Python, runs on your Mac) that picks up the actions
  Office.js can't do — install Solver, run Solver, create scenarios,
  run Analysis ToolPak, install/uninstall add-ins.
- A **project memory** layer (Supabase-backed) that remembers what you've
  written across turns, so tsifl doesn't redo work or overwrite cells you
  told it to leave alone.
- A **regression harness** (`tests/regression/`) that locks in known-good
  behavior so a prompt tweak can't silently break a case that used to work.

## Who this is for

- Financial analysts building DCFs, LBOs, comps tables, pitchbooks.
- Anyone doing multi-step work across Excel + R + PowerPoint + email who
  spends hours on copy-paste between apps.
- Teams who want AI assistance but need it to **remember** what's already
  done and **respect** the cells they've locked.

## Architecture

```
tsifulator.ai/
├── backend/                 # FastAPI — routes, Claude integration, memory
│   ├── main.py
│   ├── routes/
│   │   └── chat.py          # Main /chat endpoint + memory endpoints
│   ├── services/
│   │   ├── claude.py        # System prompt + tool-call loop
│   │   ├── project_memory.py  # Per-workbook memory (Supabase)
│   │   ├── computer_use.py  # Routes actions to desktop agent
│   │   └── memory.py        # Cross-session conversation history
│   └── requirements.txt
├── excel-addin/             # Office.js task pane — the UI
│   ├── manifest.xml
│   └── src/
│       ├── taskpane.html
│       ├── taskpane.js
│       └── taskpane.css
├── desktop-agent/           # Python daemon for Ribbon-dialog actions
│   ├── agent.py
│   └── requirements.txt
├── r-addin/                 # RStudio equivalent
├── word-addin/  powerpoint-addin/  gmail-client/  terminal-client/
├── tests/regression/        # Rubric-based regression suite
│   ├── cases/               # One directory per test case
│   ├── rubrics/             # Per-app evaluators
│   ├── run_tests.py
│   └── capture_case.py      # Turn a live run into a new test case
├── supabase_setup.sql       # One-time schema for Supabase
└── .env.example
```

## Setup

**To install tsifl as an analyst (non-developer):** see [INSTALL.md](INSTALL.md).

**To develop / contribute:**

```bash
# 1. Environment
cp .env.example .env
# Fill in ANTHROPIC_API_KEY and SUPABASE_URL / SUPABASE_KEY

# 2. Supabase
# Copy supabase_setup.sql → Supabase Dashboard → SQL Editor → Run

# 3. Backend
cd backend && pip install -r requirements.txt && uvicorn main:app --reload

# 4. Excel add-in dev server
cd excel-addin && npm install && npm start

# 5. Desktop agent (optional — only needed for Solver / Scenario Manager / ToolPak)
cd desktop-agent && pip install -r requirements.txt
export ANTHROPIC_API_KEY="..."
python3 agent.py

# 6. Regression suite
cd tests/regression && pip install -r requirements.txt && python3 run_tests.py
```

## Live services

- **Backend (production):** https://focused-solace-production-6839.up.railway.app
- **Guards / build tag:** `/chat/debug/guards`
- **Regression suite (CI):** runs on every push to `main` via
  `.github/workflows/regression.yml`; blocks deploys that break a case.

## Pricing

- **Free** — limited tasks/month, single-app
- **Pro** — $49/month — unlimited tasks, all apps, per-project memory
- **Team** — custom — shared workflow library, admin controls

## License

Proprietary. All rights reserved.
