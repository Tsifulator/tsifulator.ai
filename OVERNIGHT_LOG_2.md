# Overnight Build Log 2 — 2026-03-29

## Part 1: Fix PowerPoint & Word Add-ins — COMPLETE

### Changes Made
- **PowerPoint add-in (v2)**: Added canvas-based image rendering (`renderImageToCanvas`) matching Excel v40 pattern. Updated image preview to use CSP-safe canvas rendering. Build verified clean.
- **Word add-in (v2)**: Same canvas-based image rendering added. Build verified clean.
- Both add-ins already had:
  - Working webpack configs with HTTPS dev certs from `~/.office-addin-dev-certs/`
  - Correct manifest.xml files (PowerPoint: Host=Presentation port 3001, Word: Host=Document port 3002)
  - Full auth.js (Supabase) shared from Excel add-in
  - Full action executors for all Office.js operations
  - Image support (file picker, drag & drop, paste)
- Added `SETUP.md` sideloading instructions for both add-ins
- Both build clean with `npm run build` (0 errors)

### Files Modified
- `powerpoint-addin/src/taskpane.js` — v2, canvas rendering, build version display
- `word-addin/src/taskpane.js` — v2, canvas rendering, build version display
- `powerpoint-addin/SETUP.md` — NEW: sideloading instructions
- `word-addin/SETUP.md` — NEW: sideloading instructions

---

## Part 2: Gmail → Browser-Wide Floating Sidebar — COMPLETE

Rebuilt `/gmail-addon/` as a browser-wide Chrome extension:

### New Features
- **Works on ALL pages** (not just Gmail) — `<all_urls>` content script
- **Site detection**: Gmail, Google Sheets, Google Docs, Google Slides, any other page
- **Context-aware**: Reads different context per site:
  - Gmail: email threads, subjects, senders, inbox snippets
  - Google Sheets: sheet tabs, active cell, formula bar
  - Google Docs: document content, paragraphs, selection
  - Google Slides: slide count, current slide text
  - Any page: page text, selection, title, URL, meta description
- **Keyboard shortcut**: Ctrl+Shift+T / Cmd+Shift+T to toggle
- **Draggable**: Drag by header to reposition anywhere on screen
- **Resizable**: Drag left edge to resize width (280-600px)
- **Minimizable**: Minimize to floating action button (FAB), click to restore
- **Smooth animations**: CSS transition slide-in/slide-out
- **Site badge**: Shows current site context (Gmail, Sheets, Docs, etc.)
- **Same tsifl design system**: white + #0D5EAF blue

### Files Modified
- `gmail-addon/manifest.json` — v2, all_urls, keyboard commands
- `gmail-addon/background.js` — Handles any tab (not just Gmail)
- `gmail-addon/content.js` — Complete rewrite: site detection, multi-context, drag/resize/minimize
- `gmail-addon/sidebar.css` — Complete rewrite: drag handle, resize handle, FAB, animations

---

## Part 3: VS Code Extension — COMPLETE

Built `/vscode-extension/` as a full VS Code extension:

### Features
- **Activity bar icon**: tsifl icon in left sidebar
- **Webview sidebar**: Full chat UI with tsifl design system (adapted for VS Code themes)
- **Context awareness**: Reads current file, selection, language, visible text, open files, diagnostics, git status
- **Commands**:
  - `tsifl: Open Chat`
  - `tsifl: Explain Selected Code` (also in right-click menu)
  - `tsifl: Refactor Selected Code`
  - `tsifl: Generate Tests`
  - `tsifl: Fix Error`
- **Action execution**: insert_code, replace_selection, create_file, edit_file, run_terminal_command, open_file, show_diff
- **Auth**: Supabase via REST API (session persisted in vscode state)
- **Image support**: File picker, clipboard paste

### Files Created
- `vscode-extension/package.json` — Extension manifest with commands, views, menus
- `vscode-extension/extension.js` — TsiflSidebarProvider webview, action executor, context capture
- `vscode-extension/media/icon.svg` — Activity bar icon
- `vscode-extension/SETUP.md` — Installation instructions

---

## Part 4: Google Workspace Add-ons — COMPLETE

Built `/google-workspace-addon/` as a Google Apps Script project:

### Features
- **Works in Sheets, Docs, and Slides** — Card-based homepage + sidebar
- **Sidebar chat UI**: Same tsifl design system
- **Auth via Supabase REST** from Apps Script `UrlFetchApp`
- **Context capture**:
  - Sheets: Active cell, data range, formulas, sheet tabs, selection
  - Docs: Paragraphs, headings, tables, cursor position, selection
  - Slides: Slides, shapes, text, current slide, selection
- **Action execution** (server-side in Apps Script):
  - Sheets: write_cell, write_range, format_range, add_sheet, navigate_sheet, sort_range, add_chart, clear_range, freeze_panes, autofit
  - Docs: insert_text, insert_paragraph, insert_table, format_text, find_and_replace, insert_page_break, insert_header, insert_footer
  - Slides: create_slide, add_text_box, add_shape, add_table, add_image, delete_slide, set_slide_background, modify_slide
- **Image support** in sidebar (file picker, paste)
- **clasp deployment** config included

### Files Created
- `google-workspace-addon/Code.gs` — Server-side logic
- `google-workspace-addon/Sidebar.html` — Client-side chat UI
- `google-workspace-addon/appsscript.json` — Manifest with OAuth scopes
- `google-workspace-addon/.clasp.json` — Deployment config
- `google-workspace-addon/SETUP.md` — Deployment instructions

---

## Part 5: System Prompt Strengthening — COMPLETE

### Changes to `backend/services/claude.py`

1. **Updated app list**: Now lists all 11 apps (Excel, RStudio, Terminal, PowerPoint, Word, Gmail, VS Code, Google Sheets, Google Docs, Google Slides, Browser)

2. **New system prompt sections**:
   - VS Code actions (insert_code, replace_selection, create_file, edit_file, run_terminal_command, open_file, show_diff) with coding principles
   - Google Sheets actions (write_cell, write_range, format_range, etc.) with formula differences from Excel (ARRAYFORMULA, QUERY, FILTER, UNIQUE)
   - Google Docs actions (insert_text, insert_paragraph, insert_table, etc.)
   - Google Slides actions (create_slide, add_text_box, add_shape, etc.)
   - Browser context handling (page text, selection, summarize, Q&A)
   - Gmail professional templates (cold outreach, follow-up, meeting request)
   - PowerPoint professional templates (pitch deck 12-slide structure, board meeting deck, investment memo)
   - Word professional templates (financial memo, engagement letter, due diligence report)
   - VS Code workflow patterns (debug, refactor, test generation, code review)
   - Google Sheets template guidance (QUERY, ARRAYFORMULA instead of fill_down)
   - Strengthened run_shell_command restrictions for ALL app contexts

3. **Updated tool definition**: Added all VS Code and Google Workspace action types with payload schemas

4. **New context formatters**: Added `_format_context` handlers for:
   - `vscode`: workspace, file, language, selection, diagnostics, git
   - `google_sheets`: spreadsheet data, formulas, selection
   - `google_docs`: paragraphs, headings, tables, cursor, selection
   - `google_slides`: slides, shapes, text, current slide
   - `browser`: URL, title, page text, selection, meta description

5. **Test file**: Created `backend/tests/test_all_apps.py` with 40 test scenarios across 9 app contexts (5 per app)

### Test Results: 40/40 PASS (100%)
All 40 scenarios pass against the live Railway backend:

| App | Tests | Status |
|-----|-------|--------|
| PowerPoint | 5 (title slide, pitch deck, table, chart, no-shell-command) | ALL PASS |
| Word | 5 (memo, table, heading, find-replace, no-shell-command) | ALL PASS |
| Gmail | 5 (draft, reply, summarize, action items, cold outreach) | ALL PASS |
| VS Code | 5 (explain, refactor, fix error, generate tests, create file) | ALL PASS |
| Google Sheets | 5 (formula, format, chart, sort, add sheet) | ALL PASS |
| Google Docs | 5 (section, table, find-replace, header, memo) | ALL PASS |
| Google Slides | 5 (create slide, shapes, table, background, delete) | ALL PASS |
| Browser | 5 (summarize, extract, explain, selection, action items) | ALL PASS |

### Backend deployed to Railway (2 deploys)

---

## Architecture Summary

```
tsifulator.ai/
├── backend/                  — FastAPI + Claude (Railway)
├── excel-addin/              — Office.js Excel (port 3000) — EXISTING
├── powerpoint-addin/         — Office.js PowerPoint (port 3001) — FIXED
├── word-addin/               — Office.js Word (port 3002) — FIXED
├── gmail-addon/              — Chrome extension (all pages) — REBUILT
├── vscode-extension/         — VS Code extension — NEW
├── google-workspace-addon/   — Google Apps Script — NEW
├── r-addin/                  — RStudio Shiny panel — EXISTING
└── terminal-client/          — Terminal CLI — EXISTING
```

All add-ins share:
- Same tsifl design system (white + Greek flag blue #0D5EAF)
- Same Supabase auth
- Same backend (`/chat/` endpoint with app-specific context)
- Same image attachment support
- Same chat UI pattern
