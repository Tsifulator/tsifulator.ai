# Overnight Build Results — 2026-03-30

## Summary
Comprehensive update across the entire Tsifulator.ai platform covering auth persistence, chat history, UI/UX polish, cross-app features, security hardening, and new integrations.

---

## Files Changed and Why

### Backend (FastAPI)

| File | Changes |
|------|---------|
| `backend/main.py` | Added CORS lockdown (origin allowlist + regex for Chrome extensions), global error handler, transfer + calendar route registration, RStudio app mapping, notes-app special URL handling |
| `backend/routes/auth.py` | **Session persistence via Supabase** — sessions now survive Railway redeploys. In-memory remains as cache, Supabase `sessions` table is the source of truth. Added SQL migration comment. |
| `backend/routes/chat.py` | **Re-enabled chat history** with session-scoped LRU store (50 sessions, 10 messages each). Heavy-action apps (Excel, RStudio, PowerPoint) skip history to prevent action truncation. |
| `backend/routes/notes.py` | Added SQL migration comments for `notes` table. Fixed UUID generation for in-memory fallback (was using incrementing counter that reset on deploy). Improved Supabase key detection. |
| `backend/routes/transfer.py` | **NEW** — Cross-app data transfer with 5-min TTL. Endpoints: POST /transfer/store, GET /transfer/{id}, GET /transfer/pending/{app}. Enables R-to-Excel image/table transfer. |
| `backend/routes/calendar.py` | **NEW** — Calendar event CRUD. Endpoints: POST/GET/PUT/DELETE /calendar/events. In-memory store (to be replaced with Google Calendar API proxy). |
| `backend/services/usage.py` | Removed DEV_BYPASS_LIMITS always-true override. Added Supabase persistence for usage tracking. Monthly usage keyed by user_id:YYYY-MM. |
| `backend/services/claude.py` | Extracted ~300-line `_format_context()` to `services/prompts/context_formatter.py`. File reduced from 1454 to 1151 lines. |
| `backend/services/prompts/context_formatter.py` | **NEW** — Extracted context formatting for all 12+ app types. Added `calendar` app context formatter. |
| `backend/static/notes.html` | Beautiful redesign with dark mode, smooth animations, typing indicator, time-ago timestamps, improved folder UI, backend session auto-restore (checks backend FIRST). |

### Excel Add-in

| File | Changes |
|------|---------|
| `excel-addin/src/taskpane.html` | Added Notes button, quick action buttons (Summarize, Chart, Format, Totals), contextual placeholder text |
| `excel-addin/src/taskpane.css` | Full dark mode support via prefers-color-scheme. Added code block styles, typing indicator, quick actions, fadeInUp animations, auto-resize textarea |
| `excel-addin/src/taskpane.js` | Fixed `currentUser` bug (was `currentUser?.id`, should be `CURRENT_USER?.id`). Enhanced markdown renderer with code blocks + copy buttons. Added typing indicator, auto-resize, Escape to clear, quick action wiring, Notes button, create_workbook + import_image action handlers |

### Word Add-in
| File | Changes |
|------|---------|
| `word-addin/src/taskpane.html` | Added Notes button, quick actions (Proofread, Summarize, Format, Improve) |
| `word-addin/src/taskpane.css` | Full dark mode, code blocks, typing indicator, animations (matches Excel) |
| `word-addin/src/taskpane.js` | Fixed currentUser bug, enhanced markdown, typing indicator, quick actions, Notes button |

### PowerPoint Add-in
| File | Changes |
|------|---------|
| `powerpoint-addin/src/taskpane.html` | Added Notes button, quick actions (Title Slide, Agenda, Summary, Polish) |
| `powerpoint-addin/src/taskpane.css` | Full dark mode, code blocks, typing indicator, animations (matches Excel) |
| `powerpoint-addin/src/taskpane.js` | Fixed currentUser bug, enhanced markdown, typing indicator, quick actions, Notes button |

### VS Code Extension
| File | Changes |
|------|---------|
| `vscode-extension/extension.js` | Redesigned inline webview: header bar with tsifl branding + Notes button, quick actions (Explain, Fix, Refactor, Test), typing indicator, code blocks with copy buttons, auto-resize textarea, fadeInUp animations, backend session auto-restore (checks backend FIRST) |
| `vscode-extension/package.json` | No changes (already well-configured with Activity Bar icon, context menus, keyboard shortcuts) |

### Chrome Extension
| File | Changes |
|------|---------|
| `gmail-addon/sidebar.html` | Added quick action buttons (Summarize, Find info, Actions, Save note) |
| `gmail-addon/sidebar.css` | Added dark mode support, quick action styles, code block styles, typing indicator |
| `gmail-addon/panel.js` | Enhanced markdown renderer with code blocks + copy buttons, typing indicator, quick action button wiring |
| `gmail-addon/background.js` | No changes needed (action execution was already working correctly) |

### R Studio Add-in
| File | Changes |
|------|---------|
| `r-addin/R/server.R` | Added BACKEND_URL env var support. Added Notes button in header. Added dark mode CSS, typing indicator animation, code block styles, fadeInUp animation |

### Google Workspace Add-on
| File | Changes |
|------|---------|
| `google-workspace-addon/Sidebar.html` | Complete redesign: header bar with Notes button, app-type-specific quick actions (Sheets/Docs/Slides), backend session auto-restore, typing indicator, markdown rendering, dark mode, auto-resize textarea |

### Calendar Add-on (NEW)
| File | Changes |
|------|---------|
| `calendar-addon/Code.gs` | **NEW** — Google Calendar integration. Auth via Supabase, calendar context extraction (7-day lookahead), event CRUD actions (create/list/update/delete), sendChat handler |
| `calendar-addon/Sidebar.html` | **NEW** — Chat UI with quick actions (This week, New meeting, Next event), shared auth, typing indicator, dark mode |
| `calendar-addon/appsscript.json` | **NEW** — Apps Script manifest with Calendar + external request scopes |

### Shared
| File | Changes |
|------|---------|
| `shared/tsifl-theme.css` | **NEW** — Design system CSS with all color variables, dark mode variants, button/input/chat/code styles, animations, typography scale |

### Terminal Client
| File | Changes |
|------|---------|
| `terminal-client/tsifulator.py` | Added TSIFULATOR_BACKEND_URL env var support |

---

## What Works Now That Didn't Before

1. **Sessions survive Railway redeploys** — stored in Supabase `sessions` table, not just in-memory
2. **Login once, everywhere** — all apps check backend session first, auto-restore works across Excel/Word/PowerPoint/Chrome/VS Code/Notes/Google Workspace
3. **Chat history works** — session-scoped with 10 messages per session, LRU eviction at 50 sessions (skipped for heavy-action apps to prevent action truncation)
4. **Notes app opens and auto-authenticates** — checks backend session first, beautiful UI with dark mode
5. **Cross-app transfer** — R can export plots to Excel via `/transfer/store` and `/transfer/{id}` endpoints
6. **Calendar integration** — new Google Calendar add-on with event CRUD
7. **Dark mode** — all integrations support `prefers-color-scheme: dark`
8. **Code blocks with copy buttons** — all chat interfaces render code blocks properly
9. **Typing indicators** — pulsing dots animation while waiting for Claude's response
10. **Quick action buttons** — contextual shortcuts in every integration
11. **Notes accessible everywhere** — Notes button in header of every integration
12. **Usage tracking persists** — Supabase-backed with monthly reset, DEV_BYPASS removed
13. **CORS locked down** — only specific origins allowed, not wildcard
14. **Global error handler** — stack traces never exposed to clients in production
15. **UUID-based note IDs** — no more incrementing counter that resets on deploy

---

## What Still Needs Manual Setup

1. **Supabase Tables** — Run these SQL statements in your Supabase SQL Editor:
   ```sql
   -- Sessions (for auth persistence)
   CREATE TABLE IF NOT EXISTS sessions (
     id TEXT PRIMARY KEY DEFAULT 'current',
     access_token TEXT, refresh_token TEXT,
     user_id TEXT, email TEXT,
     updated_at TIMESTAMPTZ DEFAULT now()
   );

   -- Notes
   CREATE TABLE IF NOT EXISTS notes (
     id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
     user_id TEXT NOT NULL, title TEXT NOT NULL DEFAULT 'Untitled',
     content TEXT DEFAULT '', folder TEXT DEFAULT 'General',
     tags TEXT[] DEFAULT '{}',
     created_at TIMESTAMPTZ DEFAULT now(), updated_at TIMESTAMPTZ DEFAULT now()
   );
   CREATE INDEX IF NOT EXISTS idx_notes_user ON notes(user_id);

   -- Usage tracking
   CREATE TABLE IF NOT EXISTS usage (
     user_id TEXT NOT NULL, month TEXT NOT NULL,
     used INTEGER DEFAULT 0, tier TEXT DEFAULT 'starter',
     PRIMARY KEY (user_id, month)
   );
   ```

2. **Google Calendar Add-on** — Deploy via `clasp push` from `calendar-addon/` directory. Requires Google Cloud project with Calendar API enabled.

3. **Google Workspace Add-on** — Deploy via `clasp push` from `google-workspace-addon/` directory.

4. **Chrome Extension** — Load unpacked from `gmail-addon/` in Chrome > Extensions > Developer Mode.

5. **VS Code Extension** — Install the .vsix file: `code --install-extension tsifl-2.0.0.vsix`

6. **Office Add-ins** — Sideload manifests or deploy to Office Add-in store.

---

## Known Issues / Couldn't Resolve

1. **Office manifest `AutoShow`** — The `<AutoShow>` element is not widely supported in Office.js manifests. Add-ins require manual first-open via ribbon button.

2. **R `.Rprofile` auto-launch** — Not implemented (would require modifying user's global R profile which is invasive). Users should launch via RStudio Addin menu.

3. **Cross-app navigation buttons** — The `/launch-app` endpoint only works when backend runs locally (not on Railway, which is a Linux container). Browser-based integrations can open URLs but not launch desktop apps.

4. **Google Calendar API proxy** — The backend `/calendar/events` endpoint uses in-memory storage. For real calendar integration, events are managed directly via Apps Script's CalendarApp API.

5. **VS Code webview still inline** — The HTML/CSS/JS is embedded in extension.js as a template literal. Separating into files would require webpack-like bundling for the extension, which adds complexity. The current approach is functional and maintainable with the improvements made.
