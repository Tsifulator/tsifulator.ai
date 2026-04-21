-- ============================================================
-- Tsifulator.ai — Supabase Schema
-- Run this in your Supabase project: SQL Editor → New Query → Run
-- ============================================================

-- Conversation history (shared across Excel + RStudio)
create table if not exists messages (
  id          bigserial primary key,
  user_id     text        not null,
  role        text        not null check (role in ('user', 'assistant')),
  content     text        not null,
  app         text        default 'excel',
  session_id  text        default '',
  created_at  timestamptz default now()
);

-- Index for fast user history lookups
create index if not exists messages_user_id_idx on messages (user_id, created_at desc);

-- Saved model contexts (LBO, DCF, etc.) for cross-session memory
create table if not exists model_contexts (
  id          bigserial primary key,
  user_id     text        not null,
  model_type  text        not null,
  context     jsonb       not null,
  updated_at  timestamptz default now(),
  unique (user_id, model_type)
);

-- Per-workbook project memory: what cells are already correct, what's locked,
-- what's pending. Loaded on every /chat turn and injected into the system
-- prompt so the LLM knows what NOT to redo. Survives Railway redeploys.
create table if not exists project_memory_state (
  user_id     text        not null,
  workbook_id text        not null,                  -- sha256(app + sorted sheet names)[:16]
  state       jsonb       not null default '{}'::jsonb,
  updated_at  timestamptz not null default now(),
  primary key (user_id, workbook_id)
);
create index if not exists project_memory_state_updated_idx
  on project_memory_state (updated_at desc);

-- Backend writes with a trusted service_role key; app-level access rules are
-- enforced in Python, not Postgres. Disable RLS so inserts don't 42501.
alter table project_memory_state disable row level security;

-- ============================================================
-- Done. Go back to your .env and add your Supabase credentials.
-- ============================================================
