Do not ask clarifying questions. Make reasonable defaults. Start coding immediately.

You are my senior full-stack engineering agent. Build a production-oriented SaaS starter in this repository with clean architecture, small iterative commits, and runnable code after each phase.

Project context:
- Repo root: Tsifulator.ai
- OS: Windows
- Runtime: Node.js 24+
- Existing folders: server, clients/terminal, clients/office-addin, docs
- Env var already present: OPENAI_API_KEY in .env

Hard requirements:
1) Use TypeScript for backend and frontend code.
2) Backend: Node + Fastify, modular routes, health endpoint, env validation.
3) Database: Prisma + SQLite for local dev (easy switch to Postgres via env).
4) Auth: JWT-based email/password with secure password hashing.
5) Billing-ready skeleton: Stripe integration placeholders with clear TODOs.
6) API quality: Zod validation, centralized error handling, request logging.
7) Frontend (terminal client first): minimal React + Vite app with:
   - Sign up / Sign in views
   - Dashboard shell
   - API client wrapper with token handling
8) Developer experience:
   - ESLint + Prettier config
   - npm scripts for dev/build/test/lint/format
   - .env.example
   - concise README with setup/run steps
9) Keep current folder structure and place code appropriately.
10) Do not introduce unnecessary features. Keep MVP-focused.

Execution style:
- Work in phases and print each phase header.
- After each phase, list exactly what files were created/updated.
- Provide copy-paste terminal commands for each phase.
- If a command may differ on Windows, include Windows-safe command.
- Prefer deterministic defaults over optional branches.

Phase order:
Phase 1: Foundation and toolchain
Phase 2: Backend core + auth
Phase 3: Prisma schema + local persistence
Phase 4: Frontend terminal client MVP
Phase 5: Integration, docs, and runbook

Output format:
- Start with a 6-10 bullet implementation plan.
- Then provide Phase 1 immediately with code and commands.
- No questions. No placeholders like "you can" unless unavoidable.
