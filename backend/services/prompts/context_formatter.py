"""
Context Formatter — Converts raw app context into structured text for Claude.
Extracted from claude.py for maintainability.
"""


def _col_letter(idx: int) -> str:
    letters = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def format_context(context: dict) -> str:
    if not context:
        return ""

    app = context.get("app", "excel")

    if app == "excel":
        lines = ["[EXCEL WORKBOOK CONTEXT]"]
        active_sheet = context.get('sheet', 'Sheet1')
        lines.append(f"Active sheet: {active_sheet}")
        lines.append(f"Selected cell: {context.get('selected_cell', 'A1')}")

        prefs = context.get("preferences", {})
        if prefs:
            lines.append("User preferences (apply automatically):")
            for k, v in prefs.items():
                lines.append(f"  {k}: {v}")

        summaries = context.get("sheet_summaries", [])
        if summaries:
            lines.append("\n[WORKBOOK SHEET MAP — full data for every sheet]")
            for s in summaries:
                if s.get("rows", 0) == 0:
                    lines.append(f"  Sheet '{s['name']}': empty")
                    continue
                used_range_str = s.get("used_range", "")
                lines.append(f"\n  Sheet '{s['name']}' — {s.get('rows',0)} rows × {s.get('cols',0)} cols  (range: {used_range_str})")
                try:
                    addr_part = used_range_str.split("!")[1] if "!" in used_range_str else used_range_str
                    start_row = int(''.join(filter(str.isdigit, addr_part.split(":")[0])))
                except Exception:
                    start_row = 1
                preview          = s.get("preview", [])
                preview_formulas = s.get("preview_formulas", [])
                for r_idx, row in enumerate(preview):
                    actual_row = start_row + r_idx
                    non_empty = []
                    for c_idx, val in enumerate(row[:26]):
                        formula = (preview_formulas[r_idx][c_idx]
                                   if preview_formulas
                                   and r_idx < len(preview_formulas)
                                   and c_idx < len(preview_formulas[r_idx])
                                   else None)
                        display = formula if (formula and str(formula).startswith("=")) else val
                        if display not in (None, "", 0):
                            non_empty.append((c_idx, display))
                    if non_empty:
                        cells = "  ".join(f"{_col_letter(c)}{actual_row}={repr(v)}" for c, v in non_empty)
                        lines.append(f"    {cells}")

        # Named ranges (Improvement 20)
        named_ranges = context.get("named_ranges", [])
        if named_ranges:
            lines.append("\n[NAMED RANGES]")
            for nr in named_ranges[:30]:
                lines.append(f"  {nr.get('name', '')}: {nr.get('reference', '')}")

        sheet_data     = context.get("sheet_data", [])
        sheet_formulas = context.get("sheet_formulas", [])

        if sheet_data:
            lines.append(f"\n[ACTIVE SHEET: '{active_sheet}' — full data]")
            lines.append(f"Used range: {context.get('used_range', '')}")
            used_range = context.get("used_range", "")
            try:
                start_row = int(''.join(filter(str.isdigit,
                                used_range.split("!")[1].split(":")[0]
                                if "!" in used_range else used_range.split(":")[0])))
            except Exception:
                start_row = 1

            for r_idx, row in enumerate(sheet_data[:50]):
                actual_row = start_row + r_idx
                for c_idx, val in enumerate(row[:26]):
                    formula = sheet_formulas[r_idx][c_idx] if sheet_formulas and r_idx < len(sheet_formulas) and c_idx < len(sheet_formulas[r_idx]) else None
                    if formula and str(formula).startswith("="):
                        lines.append(f"  {_col_letter(c_idx)}{actual_row}: {formula}")
                    elif val not in (None, "", 0):
                        lines.append(f"  {_col_letter(c_idx)}{actual_row}: {repr(val)}")
        else:
            lines.append(f"\nActive sheet '{active_sheet}' is empty.")

    elif app == "rstudio":
        lines = ["[RSTUDIO CONTEXT]"]
        lines.append(f"R version: {context.get('r_version', 'unknown')}")
        lines.append(f"Working dir: {context.get('working_dir', '~')}")
        lines.append(f"Loaded packages: {context.get('loaded_pkgs', 'none')}")
        env_objects = context.get("env_objects", [])
        if env_objects:
            lines.append("Global environment objects:")
            for obj in env_objects:
                dim_str = f" [{obj.get('dim')}]" if obj.get('dim') else ""
                cols = ""
                if obj.get("col_names"):
                    cols = f" | columns: {obj['col_names']}"
                lines.append(f"  {obj['name']} ({obj['class']}{dim_str}){cols}: {obj.get('preview','')}")
        else:
            lines.append("Global environment is empty.")
        # Open editor tab
        open_editor = context.get("open_editor", {})
        if open_editor:
            if open_editor.get("active_file"):
                lines.append(f"Active editor tab: {open_editor['active_file']}")
            if open_editor.get("active_preview"):
                lines.append(f"Active file preview:\n{open_editor['active_preview']}")

    elif app == "terminal":
        lines = ["[TERMINAL CONTEXT]"]
        lines.append(f"Shell: {context.get('shell', 'zsh')}")
        lines.append(f"Working dir: {context.get('working_dir', '~')}")
        recent = context.get("recent_commands", [])
        if recent:
            lines.append("Recent commands: " + ", ".join(recent))
        ls_files = context.get("ls", [])
        if ls_files:
            lines.append(f"Files: {', '.join(ls_files[:15])}")

    elif app == "powerpoint":
        lines = ["[POWERPOINT CONTEXT]"]
        lines.append(f"Total slides: {context.get('total_slides', 0)}")
        current_slide = context.get("current_slide", {})
        if current_slide:
            lines.append(f"Current slide index: {current_slide.get('index', 0)}")
            lines.append(f"Layout: {current_slide.get('layout', 'unknown')}")
        slides = context.get("slides", [])
        if slides:
            lines.append("\n[SLIDE MAP]")
            for s in slides:
                lines.append(f"  Slide {s.get('index', 0)}: {s.get('title', '(no title)')}")
                shapes = s.get("shapes", [])
                for sh in shapes[:10]:
                    lines.append(f"    - {sh.get('type', 'shape')}: {sh.get('text', '')[:80]}")

    elif app == "word":
        lines = ["[WORD DOCUMENT CONTEXT]"]
        lines.append(f"Total paragraphs: {context.get('total_paragraphs', 0)}")
        lines.append(f"Total pages: {context.get('total_pages', 'unknown')}")
        selection = context.get("selection", "")
        if selection:
            lines.append(f"Selected text: {selection[:200]}")
        paragraphs = context.get("paragraphs", [])
        if paragraphs:
            lines.append("\n[DOCUMENT CONTENT]")
            for p in paragraphs[:50]:
                style = p.get("style", "Normal")
                text = p.get("text", "")
                if text.strip():
                    lines.append(f"  [{style}] {text[:120]}")
        tables = context.get("tables", [])
        if tables:
            lines.append(f"\n[TABLES: {len(tables)} found]")
            for i, t in enumerate(tables[:5]):
                lines.append(f"  Table {i+1}: {t.get('rows', 0)} rows × {t.get('columns', 0)} cols")

    elif app == "gmail":
        lines = ["[GMAIL CONTEXT]"]
        lines.append(f"Account: {context.get('email', 'connected')}")
        recent_emails = context.get("recent_emails", [])
        if recent_emails:
            lines.append("Recent emails:")
            for e in recent_emails[:5]:
                lines.append(f"  {e.get('from','')} — {e.get('subject','')}")
        current_thread = context.get("current_thread", {})
        if current_thread:
            lines.append(f"\nCurrent thread: {current_thread.get('subject', '')}")
            messages = current_thread.get("messages", [])
            for m in messages[:10]:
                lines.append(f"  From: {m.get('from', '')} — {m.get('snippet', '')[:100]}")

    elif app == "vscode":
        lines = ["[VS CODE CONTEXT]"]
        lines.append(f"Workspace: {context.get('workspace', '')}")
        lines.append(f"Current file: {context.get('current_file', 'none')}")
        lines.append(f"Language: {context.get('language', 'unknown')}")
        lines.append(f"Lines: {context.get('line_count', 0)}")
        cursor_line = context.get("cursor_line", 0)
        if cursor_line:
            lines.append(f"Cursor at line: {cursor_line}")
        framework = context.get("framework", "")
        if framework:
            lines.append(f"Framework: {framework}")
        open_files = context.get("open_files", [])
        if open_files:
            lines.append(f"Open files: {', '.join(open_files[:10])}")
        git_branch = context.get("git_branch", "")
        if git_branch:
            lines.append(f"Git branch: {git_branch}, {context.get('git_changes', 0)} uncommitted changes")
        diagnostics = context.get("diagnostics", [])
        if diagnostics:
            lines.append("\nDiagnostics (errors/warnings):")
            for d in diagnostics[:15]:
                lines.append(f"  {d.get('severity','')}: {d.get('file','').split('/')[-1]}:{d.get('line',0)} — {d.get('message','')}")
        selection = context.get("selection", "")
        if selection:
            lines.append(f"\n[SELECTED TEXT]\n{selection[:1500]}")
        visible = context.get("visible_text", "")
        if visible:
            lines.append(f"\n[VISIBLE CODE (what user is looking at)]\n{visible[:3000]}")
        file_content = context.get("file_content", "")
        if file_content and file_content != visible:
            lines.append(f"\n[FULL FILE (first 5000 chars)]\n{file_content[:5000]}")
        file_tree = context.get("file_tree", [])
        if file_tree:
            tree_str = ", ".join(f.get("path", "") for f in file_tree[:30])
            lines.append(f"\n[PROJECT FILES]\n{tree_str}")

    elif app == "google_sheets":
        lines = ["[GOOGLE SHEETS CONTEXT]"]
        lines.append(f"Spreadsheet: {context.get('spreadsheet_name', '')}")
        lines.append(f"Active sheet: {context.get('sheet_name', 'Sheet1')}")
        lines.append(f"All sheets: {', '.join(context.get('all_sheets', []))}")
        lines.append(f"Active cell: {context.get('active_cell', 'A1')}")
        lines.append(f"Data range: {context.get('data_range', '')} ({context.get('row_count', 0)} rows × {context.get('col_count', 0)} cols)")
        data = context.get("data", [])
        formulas = context.get("formulas", [])
        if data:
            lines.append("\n[SHEET DATA]")
            for r_idx, row in enumerate(data[:40]):
                for c_idx, val in enumerate(row[:20]):
                    formula = formulas[r_idx][c_idx] if formulas and r_idx < len(formulas) and c_idx < len(formulas[r_idx]) else ""
                    display = formula if formula else val
                    if display not in (None, "", 0):
                        lines.append(f"  {_col_letter(c_idx)}{r_idx+1}: {repr(display)}")
        sel_vals = context.get("selection_values", [])
        if sel_vals:
            lines.append(f"\nSelection values ({context.get('active_cell', '')}):")
            for row in sel_vals[:10]:
                lines.append(f"  {row}")

    elif app == "google_docs":
        lines = ["[GOOGLE DOCS CONTEXT]"]
        lines.append(f"Document: {context.get('document_name', '')}")
        lines.append(f"Paragraphs: {context.get('paragraph_count', 0)}")
        selection = context.get("selection", "")
        if selection:
            lines.append(f"Selected text: {selection[:500]}")
        cursor_text = context.get("cursor_text", "")
        if cursor_text:
            lines.append(f"Cursor at: {cursor_text}")
        paragraphs = context.get("paragraphs", [])
        if paragraphs:
            lines.append("\n[DOCUMENT CONTENT]")
            for p in paragraphs[:40]:
                if p.get("type") == "table":
                    lines.append(f"  [TABLE: {p.get('rows',0)}×{p.get('cols',0)}]")
                else:
                    heading = p.get("heading", "NORMAL")
                    text = p.get("text", "")
                    if text.strip():
                        lines.append(f"  [{heading}] {text[:120]}")

    elif app == "google_slides":
        lines = ["[GOOGLE SLIDES CONTEXT]"]
        lines.append(f"Presentation: {context.get('presentation_name', '')}")
        lines.append(f"Total slides: {context.get('slide_count', 0)}")
        lines.append(f"Current slide: {context.get('current_slide_index', 0)}")
        selection = context.get("selection", "")
        if selection:
            lines.append(f"Selected text: {selection[:300]}")
        slides = context.get("slides", [])
        if slides:
            lines.append("\n[SLIDE MAP]")
            for s in slides:
                lines.append(f"  Slide {s.get('index', 0)}: {s.get('title', '(no title)')}")
                for sh in s.get("shapes", [])[:8]:
                    lines.append(f"    - {sh.get('type', 'shape')}: {sh.get('text', '')[:60]}")

    elif app == "calendar":
        lines = ["[CALENDAR CONTEXT]"]
        lines.append(f"Calendar: {context.get('calendar_name', 'Default')}")
        lines.append(f"Timezone: {context.get('timezone', 'UTC')}")
        lines.append(f"Current time: {context.get('current_time', '')}")
        events = context.get("upcoming_events", [])
        if events:
            lines.append(f"\n[UPCOMING EVENTS ({len(events)})]")
            for e in events:
                lines.append(f"  {e.get('title', 'No title')} | {e.get('start', '')} - {e.get('end', '')}")
                if e.get('description'):
                    lines.append(f"    Description: {e['description'][:100]}")
                if e.get('guests'):
                    lines.append(f"    Guests: {', '.join(e['guests'][:5])}")
        else:
            lines.append("No upcoming events.")

    elif app == "notes":
        lines = ["[NOTES CONTEXT]"]
        lines.append(f"Note title: {context.get('note_title', 'Untitled')}")
        note_content = context.get("note_content", "")
        if note_content:
            lines.append(f"\n[NOTE CONTENT]\n{note_content[:10000]}")
        else:
            lines.append("Note is empty.")

    elif app == "browser":
        lines = ["[BROWSER CONTEXT]"]
        lines.append(f"URL: {context.get('url', '')}")
        lines.append(f"Title: {context.get('title', '')}")
        meta = context.get("meta_description", "")
        if meta:
            lines.append(f"Description: {meta}")
        thread_subject = context.get("thread_subject", "")
        if thread_subject:
            lines.append(f"Email thread: {thread_subject}")
        messages = context.get("messages", [])
        if messages:
            lines.append("Thread messages:")
            for m in messages[:5]:
                lines.append(f"  {m.get('sender', '')}: {m.get('snippet', '')[:200]}")
        sheet_title = context.get("sheet_title", "")
        if sheet_title:
            lines.append(f"Spreadsheet: {sheet_title}")
        doc_title = context.get("doc_title", "")
        if doc_title:
            lines.append(f"Document: {doc_title}")
        doc_content = context.get("doc_content", "")
        if doc_content:
            lines.append(f"Document content:\n{doc_content[:3000]}")
        selection = context.get("selection", "")
        if selection:
            lines.append(f"\nSelected text:\n{selection[:1500]}")
        full_page_text = context.get("full_page_text", "")
        if full_page_text:
            lines.append(f"\n[FULL PAGE TEXT FOR SUMMARIZATION]\n{full_page_text[:12000]}")
        elif not selection:
            page_text = context.get("page_text", "")
            if page_text:
                lines.append(f"\nPage content:\n{page_text[:2500]}")

    else:
        return ""

    return "\n".join(lines)
