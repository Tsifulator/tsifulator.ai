# OVERNIGHT MASTER PROMPT

Copy everything below the `---` line into terminal and let it run overnight.

---

You are working on the Tsifulator.ai codebase at `/Users/nicholastsiflikiotis/tsifulator.ai/`. This is an agentic AI sandbox for financial analysts — tsifl is embedded inside Excel, Word, PowerPoint, RStudio, VS Code, Chrome, Google Workspace (Sheets/Docs/Slides), Terminal, Gmail, and a Notes app. All integrations connect to a FastAPI backend deployed on Railway at `https://focused-solace-production-6839.up.railway.app`.

You have the ENTIRE NIGHT to work through this list. Be methodical. Work through each section completely before moving on. Test your changes mentally. Read every file before editing. DO NOT skip anything. DO NOT leave TODOs. DO NOT break existing functionality. After EACH major section, run `cd /Users/nicholastsiflikiotis/tsifulator.ai/backend && railway up` to deploy.

## CODEBASE STRUCTURE
```
tsifulator.ai/
├── backend/           → FastAPI (main.py, routes/, services/)
├── excel-addin/       → Office.js Excel add-in (src/taskpane.js, auth.js)
├── word-addin/        → Office.js Word add-in
├── powerpoint-addin/  → Office.js PowerPoint add-in
├── vscode-extension/  → VS Code extension (extension.js)
├── r-addin/           → RStudio Shiny add-in (R/server.R)
├── gmail-addon/       → Chrome extension (manifest.json, panel.js, background.js, content.js, sidebar.html)
├── google-workspace-addon/ → Google Apps Script (Code.gs, Sidebar.html)
├── notes-app/         → Standalone notes (index.html)
├── terminal-client/   → CLI client (tsifulator.py)
├── brand/             → Logos/icons
└── shared/            → (empty, to be used)
```

---

# ═══════════════════════════════════════════════════════
# PHASE 1: FIX BROKEN THINGS FIRST
# ═══════════════════════════════════════════════════════

## 1A. FIX THE NOTES APP (NOT OPENING)

The Notes app is served at `GET /notes-app` from `backend/static/notes.html`. It exists but the user can't open it. Fix:

1. Read `backend/static/notes.html` and `backend/routes/notes.py` fully.
2. The notes.html has its own login form but it should AUTO-RESTORE the session from `/auth/get-session` endpoint on page load — so if the user is logged in from ANY tsifl app, notes opens directly without asking for login.
3. Add this logic at the TOP of the notes.html `<script>`:
   - On page load, fetch `GET /auth/get-session` from the backend
   - If tokens exist, use them to call `supabase.auth.setSession({access_token, refresh_token})`
   - If that succeeds, skip the login screen and show the app immediately
   - Only show login if no stored session AND setSession fails
4. Make sure the Supabase `notes` table exists. Add a SQL migration comment in notes.py showing the required table schema:
   ```sql
   CREATE TABLE IF NOT EXISTS notes (
     id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
     user_id TEXT NOT NULL,
     title TEXT NOT NULL DEFAULT 'Untitled',
     content TEXT DEFAULT '',
     folder TEXT DEFAULT 'General',
     tags TEXT[] DEFAULT '{}',
     created_at TIMESTAMPTZ DEFAULT now(),
     updated_at TIMESTAMPTZ DEFAULT now()
   );
   CREATE INDEX idx_notes_user ON notes(user_id);
   ```
5. Add a fallback: if Supabase table doesn't exist, use in-memory store but make it actually work properly.
6. Make the notes app BEAUTIFUL. Clean UI with smooth transitions, folder sidebar, note editor, AI action buttons (summarize, expand, rewrite, extract action items, ask AI).
7. Make notes accessible from EVERY other tsifl integration — add a "Notes" button in the header of Excel, Word, PowerPoint, VS Code, Chrome extension, and R Studio add-ins that opens the notes URL.

## 1B. FIX SESSION PERSISTENCE (LOST ON RAILWAY REDEPLOY)

The biggest issue: sessions are stored in-memory (`_session_store` dict in `auth.py`) and wiped every Railway deploy. Fix:

1. Read `backend/routes/auth.py` fully.
2. Store sessions in the Supabase `sessions` table instead of in-memory:
   ```sql
   CREATE TABLE IF NOT EXISTS sessions (
     id TEXT PRIMARY KEY DEFAULT 'current',
     access_token TEXT,
     refresh_token TEXT,
     user_id TEXT,
     email TEXT,
     updated_at TIMESTAMPTZ DEFAULT now()
   );
   ```
3. Update `set-session` to write to Supabase table.
4. Update `get-session` to read from Supabase table.
5. Keep in-memory as a CACHE layer (check memory first, fall back to Supabase).
6. This way sessions survive container restarts.

## 1C. FIX SHARED AUTH (STILL REQUIRING LOGIN EVERYWHERE)

1. Read ALL auth.js files: `excel-addin/src/auth.js`, `word-addin/src/auth.js`, `powerpoint-addin/src/auth.js`
2. In EACH add-in's `getCurrentUser()`:
   - First check local Supabase session
   - If exists, sync to backend (`/auth/set-session`) and return user
   - If NOT exists, try `restoreSessionFromBackend()` → calls `/auth/get-session`, gets tokens, calls `supabase.auth.setSession()`
   - If THAT fails, show login screen
3. When user signs in on ANY add-in, IMMEDIATELY sync to backend so all others pick it up.
4. Do the same for the Chrome extension (`gmail-addon/panel.js`) — restore from `/auth/get-session` on load.
5. Do the same for VS Code extension (`vscode-extension/extension.js`).
6. Do the same for Notes app (`backend/static/notes.html`).
7. Do the same for Google Workspace (`google-workspace-addon/Sidebar.html`).
8. The user should log in ONCE in ANY app and be logged in EVERYWHERE.

## 1D. FIX CHAT HISTORY (COMPLETELY DISABLED)

1. Read `backend/routes/chat.py` — line 64 has `history = []` hardcoded.
2. Re-enable conversation history but with SESSION SCOPING:
   - Each app session gets its own conversation thread
   - Use session_id from the request (or generate one per app instance)
   - Keep last 10 messages per session (not unlimited — prevents token bloat)
   - Store in memory with a max of 50 sessions (LRU eviction)
3. Read `backend/services/memory.py` and make `get_recent_history()` actually return data.
4. Pass history to `get_claude_response()` properly.

## 1E. FIX CHROME EXTENSION (UNRELIABLE)

1. Read ALL files in `gmail-addon/`: manifest.json, background.js, panel.js, content.js, sidebar.html, sidebar.css.
2. Fix action execution — when Claude returns `open_url`, `search_web`, etc., they MUST actually execute:
   - In `panel.js`: when receiving actions from backend, send them to `background.js` via `chrome.runtime.sendMessage()`
   - In `background.js`: handle each action type properly:
     - `open_url` → `chrome.tabs.create({url})`
     - `search_web` → `chrome.tabs.create({url: 'https://www.google.com/search?q=' + query})`
     - `open_url_current_tab` → `chrome.tabs.update(tabId, {url})`
   - Confirm the action executed and send status back to panel
3. Fix the content script double-injection guard.
4. Make sure clicking the extension icon ALWAYS opens the side panel (never a random page).
5. The extension should work on ALL websites, not just Gmail.

---

# ═══════════════════════════════════════════════════════
# PHASE 2: PERFECT EACH INTEGRATION
# ═══════════════════════════════════════════════════════

## 2A. EXCEL ADD-IN — PERFECT IT

1. Read `excel-addin/src/taskpane.js`, `taskpane.html`, `taskpane.css`, `auth.js` fully.
2. Make tsifl THE FIRST THING you see when opening Excel:
   - The task pane should auto-open on Excel launch (add `autoOpen` in manifest.xml if possible)
   - If can't auto-open, make the ribbon button prominent with the tsifl icon
3. Add a "NEW WINDOW" capability:
   - When the user asks tsifl to create a new workbook/spreadsheet while no workbook is open, tsifl should handle it gracefully
   - Add an action type `create_workbook` that creates a new blank workbook
   - In taskpane.js, handle this action with `Excel.createWorkbook()`
4. Add markdown rendering in chat messages — use a lightweight markdown-to-HTML converter (write one inline, ~50 lines):
   - Bold: **text** → <strong>text</strong>
   - Code blocks: ```code``` → <pre><code>code</code></pre>
   - Inline code: `code` → <code>code</code>
   - Bullet points: - item → <li>item</li>
   - Headers: ## text → <h3>text</h3>
5. Add a COPY button on code blocks in chat.
6. Add keyboard shortcuts:
   - Cmd+Enter / Ctrl+Enter to send message
   - Escape to clear input
7. Add dark mode support in CSS:
   ```css
   @media (prefers-color-scheme: dark) { /* dark theme colors */ }
   ```
8. Add a Notes button in the header bar that opens the notes URL.
9. Add smooth loading animation while waiting for Claude's response (pulsing dots or skeleton).
10. Show recent conversation history in the chat panel (don't start blank every time).
11. Add the ability to import R plots/graphs INTO Excel:
    - New action type: `import_image` with payload `{image_data (base64), cell, sheet, width?, height?}`
    - In taskpane.js, handle this by inserting the image into the worksheet at the specified cell
    - Use `worksheet.shapes.addImage()` or equivalent Office.js API
    - This allows R-generated charts to be embedded directly into Excel sheets

## 2B. WORD ADD-IN — PERFECT IT

1. Read `word-addin/src/taskpane.js`, `taskpane.html`, `taskpane.css` fully.
2. Same improvements as Excel:
   - Auto-open or prominent ribbon placement
   - Markdown rendering in chat
   - Copy button on code blocks
   - Keyboard shortcuts (Cmd+Enter)
   - Dark mode CSS
   - Notes button
   - Loading animation
   - Session auto-restore
3. When no document is open, tsifl should be able to create a new one.
4. Make the chat panel feel premium — clean typography, smooth transitions, professional spacing.
5. Ensure ALL Word action types work correctly: insert_text, insert_paragraph, insert_table, format_text, etc.

## 2C. POWERPOINT ADD-IN — PERFECT IT

1. Read `powerpoint-addin/src/taskpane.js`, `taskpane.html`, `taskpane.css` fully.
2. Same improvements as Excel and Word.
3. When no presentation is open, handle gracefully (create new or show helpful message).
4. Ensure ALL PowerPoint action types work: create_slide, add_text_box, add_shape, add_table, etc.
5. Add ability to import images/charts from R into slides.

## 2D. VS CODE EXTENSION — PERFECT IT

1. Read `vscode-extension/extension.js` and `package.json` fully.
2. The inline HTML/CSS (360+ lines embedded in JS strings) is fragile. Separate into proper files:
   - Create `vscode-extension/media/sidebar.html`
   - Create `vscode-extension/media/sidebar.css`
   - Create `vscode-extension/media/sidebar.js`
   - Load them via `webview.html = fs.readFileSync(...)` with proper CSP and nonce
3. Make the extension VISIBLE immediately:
   - It should appear in the Activity Bar (left sidebar) with the tsifl icon
   - Show a welcome view when first installed
   - Add status bar item showing "tsifl" at the bottom
4. Add proper VS Code integration:
   - Right-click context menu: "Ask tsifl about this code", "Fix with tsifl", "Explain with tsifl"
   - Code lens above functions: "tsifl: Explain | Refactor | Test"
   - Error detection: when the terminal shows an error, auto-suggest "Fix with tsifl?"
5. Add keyboard shortcut: Cmd+Shift+T to toggle tsifl panel
6. Session auto-restore from backend.
7. Notes button in the sidebar header.
8. Repackage as .vsix after changes: `cd vscode-extension && npx @vscode/vsce package --no-dependencies`

## 2E. R STUDIO ADD-IN — PERFECT IT + IMPORT TO EXCEL

1. Read `r-addin/R/server.R` fully (it's long, ~650 lines).
2. The R add-in now sends loaded packages and env objects to the backend. Make sure this works perfectly.
3. Add R-to-Excel image export:
   - When R generates a plot, capture it as base64 PNG
   - Add a new action type `export_to_excel` with payload `{image_data, target_cell, target_sheet}`
   - When executed, send the image data to the Excel add-in's backend endpoint
   - Create a new backend endpoint: `POST /transfer/r-to-excel` that stores the image temporarily
   - Excel add-in can poll or receive this image and insert it
4. Add ability to create new R script windows:
   - When no editor is open, `rstudioapi::documentNew()` should still work
   - When the user says "write me a script for X", use `create_r_script` action type
   - When the user says "run this", use `run_r_code` action type
5. Add a dataset browser in the Shiny UI:
   - Show loaded data frames in the sidebar
   - Click a data frame to see `head()` preview
   - Show column names and types
6. Make the chat panel show R output inline:
   - Capture console output from `sendToConsole()` and display in chat
   - Show plots inline in the chat (as base64 images)
7. Add Notes button in the Shiny UI header.
8. Make tsifl the FIRST THING visible when opening RStudio:
   - The add-in should auto-launch as a background job when RStudio starts
   - Add to `.Rprofile`: `if (interactive()) tsifulator::start_tsifl()`
   - The Viewer pane should show the tsifl UI automatically

## 2F. GMAIL / CHROME EXTENSION — PERFECT IT

1. Read ALL files in `gmail-addon/` again after Phase 1 fixes.
2. Make the side panel premium:
   - Smooth animations on message send/receive
   - Typing indicator (three pulsing dots) while waiting
   - Message fade-in animation
   - Markdown rendering for responses
   - Code syntax highlighting for code blocks
   - Copy button on code blocks
3. Gmail-specific features:
   - Auto-detect when viewing an email → show "Summarize this email" quick action
   - "Draft reply" button when viewing an email thread
   - "Extract action items" from email threads
   - Show email context in tsifl (subject, sender, snippet)
4. Browser-wide features:
   - On any webpage: "Summarize this page" quick action button
   - "Find on page" enhanced with AI
   - Form auto-fill suggestions
5. Make keyboard shortcut work: Cmd+Shift+U to toggle side panel
6. Auth auto-restore on every panel open.
7. Notes button in the panel header.

## 2G. GOOGLE WORKSPACE — IMPLEMENT PROPERLY

1. Read `google-workspace-addon/Code.gs` and `Sidebar.html` fully.
2. The Google Workspace add-on uses Apps Script. It needs:
   - FULL sidebar UI (not just a button to open sidebar)
   - The sidebar should be the same chat UI as other platforms
   - Auth via `/auth/get-session` (restore session, no separate login)
3. Google Sheets integration:
   - Read active sheet data (cell values, formulas) and send as context
   - Execute actions: write to cells, create charts, format ranges
   - Function: `getSheetContext_()` that reads the active sheet and returns {sheet, data, selection}
   - Function: `executeAction_(action)` that performs writes/formats
4. Google Docs integration:
   - Read document text and send as context
   - Execute actions: insert text, format, find/replace
   - Function: `getDocContext_()` that reads the active document
   - Function: `executeDocAction_(action)` that performs edits
5. Google Slides integration:
   - Read current slide content
   - Execute actions: create slides, add text boxes, shapes
6. Make the sidebar appear AUTOMATICALLY when opening Sheets/Docs/Slides:
   - Use `onOpen(e)` trigger to show sidebar or at least show card
   - The homepage trigger should show a rich card with quick actions
7. Add auth restoration from backend in the Sidebar.html script.
8. Deploy instructions: update SETUP.md with clear steps for deploying via `clasp push`.

## 2H. CALENDAR INTEGRATION (NEW)

1. Create a new directory: `calendar-addon/`
2. This should work as a Google Calendar add-on (Apps Script):
   - Show upcoming events
   - Create events from natural language: "Schedule a meeting with John tomorrow at 2pm"
   - Action types: `create_event`, `list_events`, `update_event`, `delete_event`
3. Add calendar context to the backend:
   - New backend endpoint: `GET /calendar/events` and `POST /calendar/events`
   - Or handle via Google Apps Script directly using `CalendarApp`
4. Add to the system prompt in `claude.py`:
   ```
   Calendar: create_event, list_events, update_event, delete_event.
   create_event: {title, start_time, end_time, description?, attendees?[]}.
   list_events: {date?, days_ahead?}.
   ```
5. Connect it with shared auth.

---

# ═══════════════════════════════════════════════════════
# PHASE 3: CROSS-APP CAPABILITIES
# ═══════════════════════════════════════════════════════

## 3A. R-TO-EXCEL IMAGE/DATA TRANSFER

Create a transfer system so R outputs can flow into Excel:

1. Backend endpoint `POST /transfer/store`:
   - Accepts: `{from_app, to_app, data_type (image/table/text), data (base64 or JSON), metadata}`
   - Stores in memory with a unique transfer_id
   - Returns: `{transfer_id}`

2. Backend endpoint `GET /transfer/{transfer_id}`:
   - Returns the stored transfer data
   - Auto-deletes after retrieval (one-time use)

3. R add-in: After generating a plot, capture it:
   ```r
   # Save current plot as base64
   tmp <- tempfile(fileext = ".png")
   dev.copy(png, tmp, width = 800, height = 600)
   dev.off()
   img_data <- base64enc::base64encode(tmp)
   # POST to /transfer/store
   ```

4. Excel add-in: New action `import_r_output`:
   - Fetches from `/transfer/{id}`
   - If image: inserts into worksheet
   - If table: writes to range

5. Update system prompt so Claude knows about this flow.

## 3B. CROSS-APP NAVIGATION

Every tsifl integration should have quick-launch buttons for other integrations:

1. In the header/toolbar of EACH add-in, add icons/buttons:
   - 📊 Excel → opens Excel or switches to it
   - 📝 Word → opens Word
   - 📑 PowerPoint → opens PowerPoint
   - 💻 VS Code → opens VS Code
   - 📈 RStudio → opens RStudio
   - 🌐 Chrome → focuses Chrome
   - 📋 Notes → opens notes URL
   - 📅 Calendar → opens calendar
2. These buttons call the backend `/launch-app` endpoint.
3. In the Chrome extension, these open the app via `chrome.tabs.create()` or system commands.

## 3C. NEW WINDOW CREATION

When no app window is open and user asks tsifl to do something:

1. Excel: Use `Excel.createWorkbook()` to create new workbook
2. Word: Use `Word.createDocument()` to create new document
3. PowerPoint: Handle gracefully (show message to open PPT first, or use `/launch-app`)
4. R: Use `rstudioapi::documentNew()` to create new R script
5. Add these as new action types in the backend system prompt:
   - `create_workbook: {}`
   - `create_document: {}`
   - `create_r_script: {code, title?}` (already exists)

---

# ═══════════════════════════════════════════════════════
# PHASE 4: UI/UX POLISH (MAKE IT FEEL PREMIUM)
# ═══════════════════════════════════════════════════════

## 4A. CONSISTENT DESIGN SYSTEM

Create a shared CSS design system. All integrations should look and feel the same:

1. Create `shared/tsifl-theme.css` with:
   - CSS custom properties (variables) for all colors
   - Light mode and dark mode variants
   - Typography scale
   - Common component styles (buttons, inputs, chat bubbles, headers)
   - Animation keyframes (fade-in, pulse, slide-up)

2. Color palette:
   ```css
   :root {
     --tsifl-blue: #0D5EAF;
     --tsifl-blue-hover: #0A4896;
     --tsifl-blue-light: #EBF3FB;
     --tsifl-bg: #F8FAFC;
     --tsifl-surface: #FFFFFF;
     --tsifl-border: #E2E8F0;
     --tsifl-text: #1E293B;
     --tsifl-muted: #64748B;
     --tsifl-green: #16A34A;
     --tsifl-red: #DC2626;
   }
   @media (prefers-color-scheme: dark) {
     :root {
       --tsifl-bg: #0F172A;
       --tsifl-surface: #1E293B;
       --tsifl-border: #334155;
       --tsifl-text: #F1F5F9;
       --tsifl-muted: #94A3B8;
       --tsifl-blue-light: #1E3A5F;
     }
   }
   ```

3. Apply this to ALL integrations:
   - Excel, Word, PowerPoint taskpane.css
   - Chrome extension sidebar.css
   - VS Code extension sidebar.css
   - R add-in CSS in server.R
   - Notes app CSS
   - Google Workspace Sidebar.html

## 4B. CHAT UI IMPROVEMENTS (ALL PLATFORMS)

Apply these to EVERY chat interface:

1. **Markdown rendering** — write a lightweight `renderMarkdown(text)` function:
   - Bold, italic, code blocks, inline code, bullet lists, numbered lists, headers, links
   - Apply to ALL assistant messages before displaying

2. **Code block styling** — code blocks get:
   - Monospace font
   - Dark background (#1E293B)
   - Light text (#E2E8F0)
   - Copy button (top-right corner)
   - Language label if specified

3. **Typing indicator** — three pulsing dots animation while waiting for response

4. **Message animations** — messages fade in from bottom with `animation: fadeInUp 0.3s ease`

5. **Timestamps** — show "just now", "2m ago" on messages

6. **Quick action buttons** — contextual buttons above the input:
   - Excel: "Summarize data", "Create chart", "Format table"
   - Word: "Proofread", "Summarize", "Format"
   - R: "Run analysis", "Create plot", "Summary stats"
   - Browser: "Summarize page", "Find info", "Take notes"

7. **Image preview** — when user attaches an image, show thumbnail before sending

8. **Chat input improvements**:
   - Auto-resize textarea (grows as user types, up to 4 lines)
   - Cmd+Enter to send
   - Shift+Enter for new line
   - Placeholder text that's contextual: "Ask tsifl about your spreadsheet..." / "Ask about this code..."

## 4C. MAKE TSIFL UNMISSABLE IN EVERY APP

The user said "make sure that all are so easily accessible in their workspace that you can't live without them like it's the first thing you notice when you open each app."

1. **Excel/Word/PowerPoint**:
   - Set `<DesktopSettings><SourceLocation>` in manifest.xml to auto-show
   - The ribbon button should be in the HOME tab (not a hidden custom tab)
   - Icon should be crisp and professional
   - Add `<AutoShow>` element in manifest.xml if supported

2. **VS Code**:
   - Activity Bar icon (left sidebar) — ALWAYS visible
   - Status bar indicator at the bottom: "tsifl ✓" when connected
   - Welcome tab on first install showing features
   - Register `onStartupFinished` activation event

3. **Chrome**:
   - Side panel opens on icon click (already set up)
   - Badge on extension icon showing notification count
   - Keyboard shortcut (Cmd+Shift+U)
   - Pin reminder on first install (install.html page)

4. **RStudio**:
   - Auto-launch via `.Rprofile`
   - Viewer pane shows tsifl
   - Addin menu entry for manual launch

5. **Google Workspace**:
   - Sidebar card appears on every Sheets/Docs/Slides open
   - Use homepage trigger to show tsifl card automatically
   - One-click to open full sidebar

---

# ═══════════════════════════════════════════════════════
# PHASE 5: SECURITY & RELIABILITY
# ═══════════════════════════════════════════════════════

## 5A. CORS LOCKDOWN

In `backend/main.py`, change:
```python
allow_origins=["*"]
```
To:
```python
allow_origins=[
    "https://localhost:3000",
    "https://localhost:3001",
    "https://localhost:3002",
    "https://focused-solace-production-6839.up.railway.app",
    "chrome-extension://*",
    "null",  # for Office add-ins
]
```

## 5B. RATE LIMITING

In `backend/services/usage.py`:
1. Remove the `DEV_BYPASS_LIMITS` override
2. Set sensible defaults: 500 tasks/month for free tier
3. Actually track usage in Supabase (not in-memory)
4. Add a `POST /usage/reset` admin endpoint for testing

## 5C. BACKEND AUTH MIDDLEWARE

Add JWT validation middleware:
1. Every endpoint (except `/`, `/auth/*`) should verify the user's Supabase JWT
2. Extract user_id from the JWT instead of trusting the request body
3. This prevents unauthorized access to chat, notes, gmail endpoints

## 5D. ERROR HANDLING

Add global error handler in `main.py`:
```python
@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    return JSONResponse(status_code=500, content={"error": "Internal server error"})
```

Add try/except in every route handler. Never expose stack traces to clients.

---

# ═══════════════════════════════════════════════════════
# PHASE 6: BACKEND IMPROVEMENTS
# ═══════════════════════════════════════════════════════

## 6A. SPLIT claude.py INTO MODULES

The file is 1,400+ lines. Split into:
- `services/claude.py` — main `get_claude_response()` function only
- `services/prompts/system_prompt.py` — the SYSTEM_PROMPT string
- `services/prompts/tools.py` — the TOOLS definition
- `services/prompts/context_formatter.py` — the `_format_context()` function

Import them in claude.py. This makes the codebase maintainable.

## 6B. ADD TRANSFER ROUTES

Create `backend/routes/transfer.py`:
- `POST /transfer/store` — store data for cross-app transfer
- `GET /transfer/{transfer_id}` — retrieve and delete transfer data
- In-memory store with 5-minute TTL

## 6C. ADD CALENDAR ROUTES

Create `backend/routes/calendar.py`:
- `POST /calendar/events` — create event (proxy to Google Calendar API)
- `GET /calendar/events` — list upcoming events
- `PUT /calendar/events/{id}` — update event
- `DELETE /calendar/events/{id}` — delete event

Register in main.py: `app.include_router(calendar.router, prefix="/calendar")`

## 6D. HARDCODED URL CLEANUP

Replace ALL hardcoded backend URLs with environment variable:
1. In `r-addin/R/server.R`: read from env var `TSIFULATOR_BACKEND_URL` or `~/.tsifulator_config`
2. In `terminal-client/tsifulator.py`: read from env var
3. In `start.sh`: already uses a variable, make sure it's consistent
4. In `google-workspace-addon/Code.gs`: use PropertiesService for configuration

---

# ═══════════════════════════════════════════════════════
# PHASE 7: DEPLOY & VERIFY
# ═══════════════════════════════════════════════════════

After ALL changes are complete:

1. Deploy backend: `cd /Users/nicholastsiflikiotis/tsifulator.ai/backend && railway up`
2. Rebuild Excel add-in: `cd /Users/nicholastsiflikiotis/tsifulator.ai/excel-addin && npx webpack --mode production`
3. Rebuild Word add-in: `cd /Users/nicholastsiflikiotis/tsifulator.ai/word-addin && npx webpack --mode production`
4. Rebuild PowerPoint add-in: `cd /Users/nicholastsiflikiotis/tsifulator.ai/powerpoint-addin && npx webpack --mode production`
5. Repackage VS Code extension: `cd /Users/nicholastsiflikiotis/tsifulator.ai/vscode-extension && npx @vscode/vsce package --no-dependencies`
6. Git commit all changes with a clear message

Write a summary in `/Users/nicholastsiflikiotis/tsifulator.ai/OVERNIGHT_RESULTS.md` listing:
- Every file you changed and why
- What works now that didn't before
- What still needs manual setup (e.g., Supabase table creation, Google Cloud project, clasp push)
- Any issues you couldn't resolve and why

---

# CRITICAL REMINDERS
- READ every file BEFORE editing it
- DO NOT delete existing functionality
- DO NOT leave placeholder/TODO code — implement fully
- DO NOT skip the deploy step after each phase
- DO NOT create unnecessary new files — edit existing ones when possible
- Use the EXACT same design language (colors, fonts, spacing) across all integrations
- Test logic mentally — if you write a fetch to `/auth/get-session`, make sure that endpoint actually returns what you expect
- The backend is at `https://focused-solace-production-6839.up.railway.app`
- Supabase URL: read from .env file
- ALL CSS should support both light and dark mode
- ALL chat UIs should have markdown rendering, code copy buttons, typing indicators, Cmd+Enter to send
- The user should NEVER have to log in more than once across all platforms
