# OpenClaw First Task (Paste this directly)

Do not ask clarifying questions. Make reasonable defaults. Start coding immediately.

## Goal
Scaffold an MVP SaaS foundation in this repository with a TypeScript Fastify backend and a minimal React terminal client.

## Deliverables
1. Backend scaffold in server/ with:
   - Fastify app entrypoint
   - /health route
   - env loader + validation (Zod)
   - error handler and logging setup
2. Database setup with Prisma in server/:
   - initial schema with User model
   - migration for local SQLite
3. Auth endpoints:
   - POST /auth/signup
   - POST /auth/signin
   - JWT issuance and protected test route
4. Frontend app in clients/terminal/:
   - Vite + React + TypeScript
   - Sign in page and dashboard shell
   - API service with token persistence
5. Repo quality:
   - ESLint + Prettier config
   - .env.example (no secrets)
   - README quickstart (Windows-safe commands)

## Constraints
- Use npm only.
- Keep dependencies minimal.
- Keep naming clear and conventional.
- No extra features beyond MVP scope.
- Ensure `npm run dev` works for backend and terminal client.

## Acceptance checks
- Backend starts and responds on /health.
- Signup/signin flow returns JWT and protected route works.
- Frontend can sign in and render dashboard shell.
- Lint passes.

## Output required from agent
- A phase-by-phase change log.
- Exact commands run.
- Final "run these 3 commands" section for local startup.
