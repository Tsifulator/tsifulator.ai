# Overnight Log 3 — 2026-03-29

## Phase 1: Fix Chrome Extension (gmail-addon/)

### Changes Made:

1. **manifest.json (v3.1.0)**
   - Added `content_scripts` declaration for auto-injection on all http/https pages
   - Added `install.html` to web accessible resources
   - Bumped version to 3.1.0

2. **background.js — Complete Rewrite**
   - Rewrote all message handlers to use async/await with Promise-based responses
   - Fixed keyboard shortcut: uses `chrome.windows.getCurrent()` instead of relying on `tab` parameter
   - Added `onInstalled` listener to re-register panel behavior on extension update
   - Added `get_page_text` message handler for full page text extraction (summarization)
   - Added `injectContentScript()` and `sendTabMessage()` helpers

3. **panel.js — Improved Reliability & Features**
   - All browser/DOM actions now route through background.js (no direct `chrome.tabs` calls)
   - Added page summarization with `isSummarizationRequest()` detection
   - Added `launch_app` action execution
   - Added `renderMarkdown()` for assistant message formatting
   - Proper `chrome.runtime.lastError` handling

4. **content.js — Enhanced Page Capture**
   - Added `getFullPageText()` — extracts up to 15K chars of structured text
   - Added `extractStructuredText()` — preserves headings, lists, blockquotes
   - Smart content element detection (article, main, .post-content, etc.)

5. **install.html — New Install Page**
   - Branded install guide with step-by-step instructions

---

## Phase 2: Fix VS Code Extension (vscode-extension/)

### Changes Made:

1. **extension.js — Full Rewrite (v1.1.0)**
   - Added `ensurePanelAndSend()` — context commands now auto-open panel if not open, then send prompt
   - All action execution sends detailed `actionComplete` messages with operation descriptions
   - `create_file` and `edit_file` now resolve relative paths against workspace root
   - `replace_selection` falls back to insert at cursor when no selection exists
   - Added `launch_app` action for cross-app capability
   - Added sign-out (double-click user bar)
   - Added Enter key for password field to trigger sign in
   - Added `renderMarkdown()` for assistant messages in webview
   - Fixed `fixError` now appears in editor right-click context menu

2. **package.json (v1.1.0)**
   - Added `fixError` to editor context menu
   - Bumped version

3. **Built tsifl-1.1.0.vsix** — ready for install

---

## Phase 3: Perfect PowerPoint and Word Add-ins

### Changes Made:

1. **PowerPoint taskpane.js**
   - Implemented `reorder_slides` action (with limitation note)
   - Implemented `apply_theme` action
   - Added `launch_app` action for cross-app capability
   - Added `renderMarkdown()` for assistant message formatting

2. **Word taskpane.js**
   - Improved `insert_table_of_contents` — adds TOC heading + update instruction + page break
   - Added `launch_app` action for cross-app capability
   - Added `renderMarkdown()` for assistant message formatting

3. **Excel taskpane.js**
   - Added `launch_app` action for cross-app capability

---

## Phase 4: Cross-App Launch Capability

### Changes Made:

1. **Backend main.py**
   - Added `POST /launch-app` endpoint
   - Supports: Excel, PowerPoint, Word, VS Code, Notes, Terminal, Safari, Chrome, Finder
   - Uses macOS `open -a` commands via subprocess

2. **All add-ins updated** — Chrome, VS Code, Excel, PowerPoint, Word all have `launch_app` action

3. **Backend claude.py** — Added `launch_app` to tool definition and browser/cross-app action types

---

## Phase 5: Apple Notes Integration

### Changes Made:

1. **notes-app/index.html — Complete Notes Web App**
   - Clean, Apple Notes-inspired UI with sidebar + editor layout
   - Folder navigation (General + custom folders)
   - Note search across titles and content
   - Auto-save on typing (1-second debounce)
   - AI integration: Summarize, Extract Action Items, free-form AI questions
   - Supabase auth with shared session sync
   - Sign out (double-click user bar)
   - Full CRUD: create, read, update, delete notes

2. **backend/routes/notes.py — Notes API**
   - GET `/notes/` — list notes (with folder/search filters)
   - GET `/notes/{id}` — get single note
   - POST `/notes/` — create note
   - PUT `/notes/{id}` — update note
   - DELETE `/notes/{id}` — delete note
   - GET `/notes/folders/list` — list all folders
   - In-memory fallback when Supabase unavailable

3. **Backend main.py** — Registered notes router, serves notes-app at `/notes-app`

4. **Backend claude.py** — Added notes context formatting, notes prompt section, text-only response mode for notes

---

## Phase 6: Continuous Prompt Strengthening

### Changes Made:

1. **Gmail prompts** — Added email writing rules: subject line format, body structure, word limits, formality matching, security guidance

2. **VS Code prompts** — Added edge cases: no selection fallback, proactive diagnostics, path resolution, language-specific rules (Python PEP 8, TypeScript strict, React hooks), import preservation, test edge cases

3. **Browser prompts** — Added edge cases per site type: Google Suite, news sites, financial sites, code repos, content validation, URL safety

4. **PowerPoint prompts** — Added precise positioning values, color system (#0D5EAF primary, #1E293B titles), table styling, chart requirements

5. **Word prompts** — Added margin values (72pt = 1 inch), font recommendations, line spacing, date formats, paragraph insertion best practices

6. **Google Sheets prompts** — Added SPARKLINE, GOOGLEFINANCE, IMPORTDATA functions, formatting edge cases, 1-indexed rows note

7. **Notes prompts** — Complete prompt section for notes context: summarization format, action item extraction format, content generation guidance

8. **Tool definition** — Added browser actions (open_url, search_web, click_element, etc.) and launch_app to type descriptions

---

## All Changes Summary

### Files Modified:
- `gmail-addon/manifest.json` — v3.1.0, content_scripts, web_accessible_resources
- `gmail-addon/background.js` — complete rewrite, async/await
- `gmail-addon/panel.js` — all actions via background, summarization, markdown
- `gmail-addon/content.js` — full page text extraction, structured text
- `vscode-extension/extension.js` — auto-open, relative paths, launch_app, sign-out
- `vscode-extension/package.json` — v1.1.0, fixError in context menu
- `powerpoint-addin/src/taskpane.js` — reorder, theme, launch_app, markdown
- `word-addin/src/taskpane.js` — TOC fix, launch_app, markdown
- `excel-addin/src/taskpane.js` — launch_app
- `backend/main.py` — notes router, notes-app endpoint, launch-app endpoint
- `backend/services/claude.py` — prompt strengthening across all 11 app contexts
- `backend/routes/notes.py` — full notes CRUD API

### Files Created:
- `gmail-addon/install.html` — Chrome extension install page
- `notes-app/index.html` — Complete AI-powered notes web app
- `backend/routes/notes.py` — Notes API
- `backend/static/notes.html` — Notes app served from backend (for Railway)
- `vscode-extension/tsifl-1.1.0.vsix` — Built extension package

---

## Phase 7: Continuous Iteration

### Changes Made:

1. **Security fixes**
   - Switched `launch-app` from `shell=True` to list-form subprocess (eliminates shell injection risk)
   - Fixed notes API route ordering (`/folders/list` before `/{note_id}`)
   - Made `launch-app` platform-aware (macOS/Windows/Linux)

2. **Reliability improvements**
   - Chrome extension: 90s request timeout + single retry on abort
   - VS Code extension: 90s request timeout via AbortController

3. **UI polish**
   - Chrome sidebar: pulse animation for thinking state, message fade-in animations, smooth transitions
   - Excel add-in: added markdown rendering for assistant messages (consistent with all other add-ins)

4. **Deployment**
   - Notes app HTML copied to `backend/static/` for Railway serving
   - Notes app live at: `https://focused-solace-production-6839.up.railway.app/notes-app`
   - All endpoints tested and verified:
     - Health check: OK
     - Notes CRUD: OK
     - Folders listing: OK
     - Launch app: OK (returns info message on Linux)
     - Auth session: OK
     - Notes web app: OK

### Deployment Status
- Backend: Deployed to Railway (5 deployments during this session)
- GitHub: 6 commits pushed to main
- All endpoints verified working
