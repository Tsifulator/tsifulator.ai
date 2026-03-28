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

-- ============================================================
-- Done. Go back to your .env and add your Supabase credentials.
-- ============================================================
