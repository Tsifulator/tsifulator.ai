"""
Claude Service — the AI brain.
Uses Anthropic tool use for guaranteed structured action output.
No text parsing needed — actions come back as typed tool calls every time.
"""

import anthropic
import os
import json
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).parent.parent.parent / ".env", override=True)

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

# ── System Prompt ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """
You are tsifl, an elite AI financial analyst and Excel powerhouse embedded inside Excel, RStudio, Terminal, and Gmail.
You READ the user's live workbook and execute real, multi-step operations with precision.

## Excel — Full Capabilities
You can perform ANY Excel operation a power user can:
- Write values AND formulas to any cell on any sheet (use write_cell with formula field for = formulas)
- Navigate between worksheets with navigate_sheet before acting on a different sheet
- Fill formulas down/right with fill_down / fill_right
- Create named ranges with create_named_range
- Sort data ranges with sort_range
- Autofit specific columns with autofit_columns
- Copy ranges with copy_range
- Apply full formatting: bold, colors, font name/size, number formats, borders, alignment, freeze panes
- Build complete financial models: LBO, DCF, 3-statement, comps, sensitivity tables

## User Preferences
The context includes a `preferences` object with the user's remembered style choices.
ALWAYS apply these preferences automatically when formatting unless the user says otherwise.
If the user says "I prefer X" or "always use X" for any style choice, call save_preference to remember it.
Default preference fallbacks (use these if no preference is set):
- header background: "#0D5EAF", header font: white, bold
- body font: "Calibri", size 11
- currency format: "$#,##0", percent: "0.0%"
- always autofit columns after bulk writes

## Multi-Sheet Operations
When working across sheets:
1. Emit navigate_sheet FIRST, then the cell/range write actions for that sheet
2. If the task touches 3 sheets, emit 3 navigate_sheet + write sequences
3. Always navigate back to the original sheet at the end if helpful

## Formula Rules
- ALWAYS use write_cell with `formula` field (not `value`) for Excel formulas
- Named ranges: create_named_range first, then reference the name in formulas
- For fill_down: write the formula in the source cell first, THEN use fill_down for the full range including that source cell
- For fill_right: write the formula in the leftmost cell first, THEN use fill_right to extend it across all columns
- Cross-sheet references use the standard Excel syntax: SheetName!CellRef
- For formulas that span multiple columns (e.g. DAVERAGE for Comfort/Fit/Style), always fill_right after writing the first column

## Named Range Rules — CRITICAL
- create_named_range MUST always include the `sheet` field explicitly — never omit it
- Example: { "type": "create_named_range", "payload": { "name": "Survey", "range": "A4:G40", "sheet": "Satisfaction Survey" } }
- The sheet field tells the add-in WHICH sheet the range lives on — without it the named range will point to the wrong data
- Always create named ranges BEFORE any formulas that reference them

## DAVERAGE / Database Function Pattern
When building DAVERAGE formulas across multiple product rows and multiple metric columns:
1. First navigate to the Criteria sheet and write ALL criteria filter values (the wildcard/product names below each header)
   - Example: if A1="Product" header, write the filter value in A2 (e.g. "rug*" or "Rugged Hiking Boots")
   - Do this for every criteria block before touching Average Ratings
2. Create the named range for the database (e.g. Survey → 'Satisfaction Survey'!A4:G40) with explicit sheet field
3. Navigate to Average Ratings and write the DAVERAGE formula in the first column (e.g. B5)
4. Use fill_right to extend the formula across all metric columns (e.g. B5:D5) — the column reference like B$4 will adjust to C$4, D$4
5. Use fill_down to extend all metric columns down for remaining product rows (e.g. B5:D9)
6. Then write the AVERAGE and IFS formulas for each row

## SUMIFS / Lookup Pattern
For SUMIFS lookups in a side table:
- Write each SUMIFS formula explicitly for every lookup row — do not leave any lookup cell empty
- Use write_cell for each lookup cell individually — never skip a row
- Example for Inventory In Stock: write_cell {cell:"L13", formula:"=SUMIFS(E$4:E$50,B$4:B$50,J13,C$4:C$50,K13)", sheet:"Inventory"}
- Then write_cell {cell:"L14", formula:"=SUMIFS(E$4:E$50,B$4:B$50,J14,C$4:C$50,K14)", sheet:"Inventory"}

## Email / Placeholder Data Pattern
If email addresses or other contact data is missing, generate realistic placeholder values:
- Format: first initial + last name + @company.com (e.g. vbowman@wearever.com)
- Always write ALL rows — never leave a data column partially filled

## Date / Time Formula Pattern
For date arithmetic columns (e.g. Days in Transit, Arrival Day):
- Days column: use write_column {start:"D5", end:"D40", formula:"=C5-B5", number_format:"0", sheet:"Shipment Times"}
- Day-of-week column: use write_column {start:"E5", end:"E40", formula:"=TEXT(C5,\"dddd\")", sheet:"Shipment Times"}
- NEVER leave these columns empty — emit write_column for BOTH columns every time

## Financial Model Guidelines
- Header rows: color "#0D5EAF" background, font_color "white", bold true, font_size 11
- Number format "$#,##0" for currency, "0.0%" for percentages, "0.00x" for multiples
- Years across columns (B, C, D...), line items down rows (A column)
- Always call autofit_columns or autofit as the final action
- Freeze first row/column when building large models

## Completeness Rule — CRITICAL
- For every column or row that is part of the task, emit ALL required actions — never leave a column half-finished
- If the task says "Days column" and "Arrival Day column", BOTH must be written
- If the task says "Comfort, Fit, Style" averages, ALL THREE columns must be filled — not just Comfort
- If a lookup table has 2 rows, write formulas for BOTH rows
- Before finishing, mentally scan every sheet and column in the task — emit actions for anything still empty

## RStudio
- Write and execute R code in the user's console
- ALWAYS include library() calls at the top — never assume packages are loaded

## Terminal / Gmail
- Execute shell commands, write files, open URLs
- Read/search/draft/send emails

## Rules
- Questions / explanations: plain text only, no execute_actions
- ANY real change: call execute_actions with all actions as a sequence
- Reply in ONE short sentence. No bullet points. No explanations.
- Respect existing sheet structure — never overwrite headers unless asked
- Never fabricate financial data — use 0 or "TBD" as placeholders
- Never emit empty or no-op actions
"""

# ── Tool Definition ───────────────────────────────────────────────────────────

TOOLS = [
    {
        "name": "execute_actions",
        "description": (
            "Execute one or more actions in the user's active app "
            "(Excel, RStudio, Terminal, or Gmail). "
            "Always call this tool — never output JSON as plain text."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "actions": {
                    "type": "array",
                    "description": "Ordered list of actions to execute.",
                    "items": {
                        "type": "object",
                        "required": ["type", "payload"],
                        "properties": {
                            "type": {
                                "type": "string",
                                "description": (
                                    "Excel cell/range: write_cell, write_formula, write_range. "
                                    "Excel navigation: navigate_sheet. "
                                    "Excel formulas: fill_down, fill_right, write_column, copy_range. "
                                    "Excel structure: create_named_range, sort_range, add_sheet, clear_range, freeze_panes. "
                                    "Excel format: format_range, set_number_format, autofit, autofit_columns. "
                                    "Preferences: save_preference. "
                                    "R: run_r_code, install_package. "
                                    "Terminal: run_shell_command, write_file, open_url. "
                                    "Gmail: send_email, draft_email, search_emails."
                                )
                            },
                            "payload": {
                                "type": "object",
                                "description": (
                                    "Payloads — all Excel actions accept optional sheet: 'SheetName' to target a specific worksheet.\n"
                                    "navigate_sheet: {sheet}.\n"
                                    "write_cell: {cell, value?, formula?, sheet?, bold?, color?, font_color?, font_size?, font_name?, number_format?, border?}.\n"
                                    "  Use formula (not value) for any cell starting with =.\n"
                                    "write_formula: {cell, formula, sheet?, bold?, color?, font_color?}.\n"
                                    "write_range: {range, values? (2D array), formulas? (2D array), sheet?, bold?, color?, font_color?, number_format?}.\n"
                                    "  IMPORTANT: values and formulas MUST be 2D arrays, e.g. [[\"val1\"],[\"val2\"]] NOT [\"val1\",\"val2\"].\n"
                                    "  Use formulas array when cells contain = formulas.\n"
                                    "write_column: {start, end, formula?, value?, sheet?, number_format?}. PREFERRED for filling a single column.\n"
                                    "  Writes formula/value to start cell and fills it down to end cell automatically.\n"
                                    "  Example: {start:'D5', end:'D40', formula:'=C5-B5', sheet:'Shipment Times'}.\n"
                                    "  Use this instead of separate write_cell + fill_down whenever filling a column.\n"
                                    "fill_down: {range, source?, sheet?}. Copies formula in first row down. source defaults to first cell of range.\n"
                                    "fill_right: {range, source, sheet?}. Copies formula in source across range.\n"
                                    "copy_range: {from, to, sheet?}. Copies values+formulas+format.\n"
                                    "create_named_range: {name, range, sheet?}. Creates a named range (workbook-level).\n"
                                    "sort_range: {range, key_column (letter), ascending?, sheet?}.\n"
                                    "add_sheet: {name, activate?}.\n"
                                    "clear_range: {range, sheet?, clear_type?}.\n"
                                    "freeze_panes: {cell?, rows?, columns?, sheet?}.\n"
                                    "format_range: {range, sheet?, bold?, italic?, color?, font_color?, font_size?, font_name?, number_format?, h_align?, border?, wrap_text?, row_height?, col_width?}.\n"
                                    "set_number_format: {range, format, sheet?}.\n"
                                    "autofit: {sheet?}. Autofits entire used range.\n"
                                    "autofit_columns: {columns: ['A','B'], sheet?} or {column: 'F', sheet?}.\n"
                                    "save_preference: {key: value, ...}. Saves user style preference to memory. Keys: font_name, font_size, header_color, header_font_color, accent_color, number_format_currency, number_format_percent.\n"
                                    "run_r_code: {code}.\n"
                                    "install_package: {package}.\n"
                                    "run_shell_command: {command}.\n"
                                    "write_file: {path, content}.\n"
                                    "open_url: {url}.\n"
                                    "send_email: {to, subject, body}.\n"
                                    "draft_email: {to, subject, body}.\n"
                                    "search_emails: {query}."
                                )
                            }
                        }
                    }
                }
            },
            "required": ["actions"]
        }
    }
]

# ── Main Entry Point ──────────────────────────────────────────────────────────

async def get_claude_response(message: str, context: dict,
                              session_id: str, history: list = []) -> dict:
    sheet_summary = _format_context(context)
    user_content  = f"{message}\n\n{sheet_summary}" if sheet_summary else message

    # Build message thread from history
    messages = []
    for h in history:
        role    = h.get("role", "user")
        content = h.get("content", "")
        app     = h.get("app", "")
        if role == "user" and app:
            content = f"[From {app}] {content}"
        messages.append({"role": role, "content": content})

    messages.append({"role": "user", "content": user_content})

    response = client.messages.create(
        model       = "claude-sonnet-4-5",
        max_tokens  = 8192,
        system      = SYSTEM_PROMPT,
        tools       = TOOLS,
        tool_choice = {"type": "auto"},  # Call tool when acting, plain text for questions
        messages    = messages,
    )

    return _parse_tool_response(response)

# ── Response Parser ───────────────────────────────────────────────────────────

def _parse_tool_response(response) -> dict:
    """Extract reply text and actions from a tool-use response."""
    reply   = ""
    actions = []

    for block in response.content:
        if block.type == "text":
            # Claude's short reply sentence
            reply = block.text.strip()

        elif block.type == "tool_use" and block.name == "execute_actions":
            # Guaranteed structured output — no parsing needed
            tool_actions = block.input.get("actions", [])
            actions.extend(tool_actions)

    # If Claude gave no reply text, use a sensible default
    if not reply:
        reply = "Done."

    return {
        "reply":   reply,
        "action":  actions[0] if len(actions) == 1 else {},
        "actions": actions    if len(actions) > 1  else [],
    }

# ── Context Formatters ────────────────────────────────────────────────────────

def _format_context(context: dict) -> str:
    if not context:
        return ""

    app = context.get("app", "excel")

    if app == "excel":
        lines = ["[EXCEL WORKBOOK CONTEXT]"]
        active_sheet = context.get('sheet', 'Sheet1')
        lines.append(f"Active sheet: {active_sheet}")
        lines.append(f"Selected cell: {context.get('selected_cell', 'A1')}")

        # User preferences
        prefs = context.get("preferences", {})
        if prefs:
            lines.append("User preferences (apply automatically):")
            for k, v in prefs.items():
                lines.append(f"  {k}: {v}")

        # ── All-sheet summaries (schema map of the workbook) ──────────────────
        summaries = context.get("sheet_summaries", [])
        if summaries:
            lines.append("\n[WORKBOOK SHEET MAP]")
            for s in summaries:
                if s.get("rows", 0) == 0:
                    lines.append(f"  Sheet '{s['name']}': empty")
                    continue
                lines.append(f"\n  Sheet '{s['name']}' — {s.get('rows',0)} rows × {s.get('cols',0)} cols  (range: {s.get('used_range','')})")
                # Show first 5 rows as structure preview
                preview = s.get("preview", [])
                for r_idx, row in enumerate(preview[:5]):
                    non_empty = [(c_idx, val) for c_idx, val in enumerate(row[:26])
                                 if val not in (None, "", 0)]
                    if non_empty:
                        cells = "  ".join(f"{_col_letter(c)}{r_idx+1}={repr(v)}" for c, v in non_empty)
                        lines.append(f"    {cells}")

        # ── Active sheet — full data + formulas ───────────────────────────────
        sheet_data     = context.get("sheet_data", [])
        sheet_formulas = context.get("sheet_formulas", [])

        if sheet_data:
            lines.append(f"\n[ACTIVE SHEET: '{active_sheet}' — full data]")
            lines.append(f"Used range: {context.get('used_range', '')}")
            # Determine start row from used_range address for accurate row labels
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
            lines.append("Global environment:")
            for obj in env_objects:
                dim = f" [{obj.get('dim')}]" if obj.get('dim') else ""
                lines.append(f"  {obj['name']} ({obj['class']}{dim}): {obj.get('preview','')}")
        else:
            lines.append("Global environment is empty.")

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

    elif app == "gmail":
        lines = ["[GMAIL CONTEXT]"]
        lines.append(f"Account: {context.get('email', 'connected')}")
        recent_emails = context.get("recent_emails", [])
        if recent_emails:
            lines.append("Recent emails:")
            for e in recent_emails[:5]:
                lines.append(f"  {e.get('from','')} — {e.get('subject','')}")
    else:
        return ""

    return "\n".join(lines)


def _col_letter(idx: int) -> str:
    letters = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters
