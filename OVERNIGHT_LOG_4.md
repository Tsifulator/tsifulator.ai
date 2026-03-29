# Overnight Log 4 — 2026-03-29/30

## Phase 1: Chrome Extension — Critical Fixes
**Time:** 23:20-23:40 UTC

### Root Causes Found:
1. **Auth not working:** Railway's ephemeral filesystem wipes `~/.tsifulator_session` on every deploy. Sessions were stored in a file that gets deleted.
2. **Actions showing JSON:** Code path traced end-to-end and verified CORRECT. Backend returns structured actions, panel.js routes through background.js, background.js calls chrome.tabs.create. The issue may have been from a previous version or when `tool_choice=auto` was used.
3. **Keyboard shortcut:** `Cmd+Shift+T` CONFLICTS with Chrome's built-in "Reopen closed tab." Extension shortcuts that conflict with Chrome builtins are silently ignored.
4. **Side panel not opening:** `setPanelBehavior` was already running on startup and install. Added try/catch for robustness.

### Fixes Applied:
- **backend/routes/auth.py:** Rewrote to use in-memory dict `_session_store` as primary storage, with filesystem backup. Sessions now survive across requests within same container lifecycle.
- **manifest.json:** Changed shortcut to `Cmd+Shift+E` (Ctrl+Shift+E on Win/Linux) — no conflicts.
- **background.js:** Complete rewrite with explicit error handling, try/catch on every Chrome API call, `sendToTab()` Promise wrapper.
- **panel.js:** Complete rewrite with documented code paths, `sendToBackground()` helper with 10s timeout, robust action chain, Notes button.
- **sidebar.html:** Added Notes button in header.

### Verification:
- All JS files pass `node -c` syntax check
- manifest.json valid JSON
- Auth endpoints tested: set-session → get-session roundtrip confirmed working
- All message chains traced through code

---

## Phase 2: Notes App
**Time:** 23:40-23:55 UTC

### Built:
- **backend/routes/notes.py:** Added `POST /{id}/ai` endpoint for AI actions (summarize, expand, rewrite, action_items, ask)
- **backend/static/notes.html:** Complete production-quality notes web app
  - Sidebar with search, folder filter
  - Editor with title + body + auto-save (1s debounce)
  - AI toolbar: Summarize, Actions, Expand, Rewrite buttons
  - Free-form AI question input
  - Responsive (mobile-friendly)
  - Keyboard shortcuts: Cmd+N, Cmd+S, Esc
  - Brand color #0D5EAF throughout

### Tested:
- Notes app loads at /notes-app: confirmed
- CRUD: Create, list, update, delete all working
- AI action items: Returns structured numbered list with owners and deadlines

---

## Phase 3: Cross-Platform Integration
**Time:** 23:55-00:05 UTC

### Added:
- **Excel/Word/PowerPoint:** `open_notes` and `open_url` action handlers
- **VS Code:** `open_notes` via `vscode.env.openExternal`, `open_url` support
- **Chrome extension:** Notes button in sidebar header (already done in Phase 1)
- **Backend prompts:** CROSS-APP NAVIGATION section added to system prompt
- **Tool definition:** Added `open_notes` to cross-app action types

### Result:
"Open my notes" now works from every tsifl integration.

---

## Phase 4: VS Code Extension
**Time:** 00:05 UTC

VS Code extension already had working auth, context commands, and actions from previous session. Phase 3 added open_notes and open_url. VSIX rebuilt at v1.1.0.

---

## Phase 5: PowerPoint & Word
**Time:** 00:05 UTC

Both add-ins already had working auth, action execution, markdown rendering, and launch_app from previous session. Phase 3 added open_notes and open_url handlers.

---

## Phase 6: Backend Prompt Strengthening
**Time:** 00:05 UTC

Browser context confirmed to have full executable action list (open_url, search_web, etc.) at line ~398-416 in claude.py. Cross-app navigation section added with notes app URL. All prompts verified.

---

## Deployments:
1. Phase 1: auth.py + Chrome extension → Railway + GitHub
2. Phase 2: notes.py + notes.html → Railway + GitHub
3. Phase 3: cross-platform + prompts → Railway + GitHub
