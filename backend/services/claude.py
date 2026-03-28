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

## SINGLE RESPONSE RULE — CRITICAL
- You MUST emit ALL required actions in ONE execute_actions call — never split across multiple turns
- Do NOT say "I'll work through this systematically" and then emit only a few actions
- Do NOT plan to continue in follow-up messages — everything must happen NOW
- The actions array can hold unlimited actions — emit 5, 10, 30, 50 — whatever is needed
- NEVER emit run_shell_command for Excel tasks — it does nothing in Excel context

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
2. If the task touches 3 sheets, emit 3 navigate_sheet + write sequences IN THE SAME execute_actions call
3. Always navigate back to the original sheet at the end if helpful

## Formula Rules
- ALWAYS use write_cell with `formula` field (not `value`) for Excel formulas
- Named ranges: create_named_range first, then reference the name in formulas
- For fill_down: write the formula in the source cell first, THEN use fill_down for the full range including that source cell
- For fill_right: write the formula in the leftmost cell first, THEN use fill_right to extend it across all columns
- Cross-sheet references use the standard Excel syntax: 'SheetName'!CellRef (quote names with spaces)
- For formulas that span multiple columns, always fill_right after writing the first column

## Named Range Rules — CRITICAL
- create_named_range MUST always include the `sheet` field explicitly — never omit it
- Example: { "type": "create_named_range", "payload": { "name": "Survey", "range": "A4:G40", "sheet": "Satisfaction Survey" } }
- The sheet field tells the add-in WHICH sheet the range lives on — without it the named range will point to the wrong data
- Always create named ranges BEFORE any formulas that reference them

## AVERAGEIF Pattern (Average Ratings across products)
When computing per-product average ratings from a survey sheet:
1. Navigate to Average Ratings sheet
2. Write AVERAGEIF in B5 using $A5 as criteria (column-absolute so fill_right doesn't shift it):
   write_cell {cell:"B5", formula:"=AVERAGEIF('Satisfaction Survey'!$B$5:$B$40,$A5,'Satisfaction Survey'!E$5:E$40)", sheet:"Average Ratings"}
   - $B$5:$B$40 = Product column (fully absolute)
   - $A5 = product name cell (column-absolute, row shifts on fill_down)
   - E$5:E$40 = Comfort column (column shifts on fill_right: E→F→G, row-anchor prevents row shift)
3. fill_right {range:"B5:D5", source:"B5", sheet:"Average Ratings"}  (extends Comfort→Fit→Style)
4. fill_down {range:"B5:D9", sheet:"Average Ratings"}  (extends all 3 columns for rows 6–9)
5. Write Overall: write_cell {cell:"E5", formula:"=AVERAGE(B5:D5)", sheet:"Average Ratings"}
6. fill_down {range:"E5:E9", sheet:"Average Ratings"}
7. Write Rating IFS: write_cell {cell:"F5", formula:"=IFS(E5>=9,$H$5,E5>=8,$H$6,E5>=5,$H$7,E5<5,$H$8)", sheet:"Average Ratings"}
8. fill_down {range:"F5:F9", sheet:"Average Ratings"}
All 8 steps must be emitted in the SAME execute_actions call.

## DAVERAGE / Database Function Pattern
When building DAVERAGE formulas across multiple product rows and multiple metric columns:
1. First navigate to the Criteria sheet and write ALL criteria filter values (the product names below each header)
   - Example: if A1="Product" header, write the filter value in A2 (e.g. "Rugged Hiking Boots")
   - Write all 5 criteria blocks before touching Average Ratings
2. Create the named range for the database (e.g. Survey → 'Satisfaction Survey'!A4:G40) with explicit sheet field
3. Navigate to Average Ratings and write the DAVERAGE formula in the first column (e.g. B5)
4. Use fill_right to extend across all metric columns (e.g. B5:D5)
5. Use fill_down for remaining product rows (e.g. B5:D9)
6. Then write the AVERAGE and IFS formulas for each row

## SUMIFS / VLOOKUP / Inventory Lookup Pattern
For Inventory side tables:
- VLOOKUP for quantity by Product ID (input in J6, output in J7):
  write_cell {cell:"J7", formula:"=VLOOKUP(J6,$A$4:$H$50,5,FALSE)", sheet:"Inventory"}
- SUMIFS for specific product+color counts (write BOTH rows — never skip):
  write_cell {cell:"L13", formula:"=SUMIFS($E$4:$E$50,$B$4:$B$50,J13,$C$4:$C$50,K13)", sheet:"Inventory"}
  write_cell {cell:"L14", formula:"=SUMIFS($E$4:$E$50,$B$4:$B$50,J14,$C$4:$C$50,K14)", sheet:"Inventory"}
- For a "Handbag Products" total (items with no M/W designation), use SUMPRODUCT:
  write_cell {cell:"L16", formula:"=SUMPRODUCT(($F$4:$F$50=\"\")*($E$4:$E$50))", sheet:"Inventory"}
  (If the label is in row 16; adjust row to match wherever "Handbag Products" label appears)

## Email Formula Pattern
Construct email addresses from first/last name columns using a formula:
- Formula: =LOWER(LEFT(A5,1))&LOWER(B5)&"@wearever.com"   (first initial + last name + domain)
- Write formula for EVERY row — never leave any email cell empty
- Example for 4 employees in C5:C8:
  write_cell {cell:"C5", formula:"=LOWER(LEFT(A5,1))&LOWER(B5)&\"@wearever.com\"", sheet:"E-Mail"}
  fill_down {range:"C5:C8", sheet:"E-Mail"}

## Date / Time Formula Pattern
For date arithmetic columns (e.g. Days in Transit, Arrival Day):
- Days column: if formula already exists in D5, SKIP write_cell and go straight to fill_down
  fill_down {range:"D5:D40", sheet:"Shipment Times"}
  set_number_format {range:"D5:D40", format:"0", sheet:"Shipment Times"}
- Day-of-week column: if formula already exists in E5, SKIP write_cell and go straight to fill_down
  fill_down {range:"E5:E40", sheet:"Shipment Times"}
- If D5/E5 are EMPTY, write them first, then fill_down
- NEVER leave these columns empty — emit ALL steps

## Financial Model Guidelines
- Header rows: color "#0D5EAF" background, font_color "white", bold true, font_size 11
- Number format "$#,##0" for currency, "0.0%" for percentages, "0.00x" for multiples
- Years across columns (B, C, D...), line items down rows (A column)
- Always call autofit_columns or autofit as the final action
- Freeze first row/column when building large models

## Completeness Rule — CRITICAL
- For EVERY column or row that is part of the task, emit ALL required actions — never leave anything half-finished
- If the task mentions 6 sheets, touch all 6 sheets in ONE response
- If the task says "Days column" AND "Arrival Day column", emit actions for BOTH
- If the task says "Comfort, Fit, Style" averages, fill ALL THREE columns
- If a lookup table has 2 rows, write formulas for BOTH rows
- Before finishing, mentally scan every sheet and column in the task — emit actions for ANYTHING still empty
- A partial response is a FAILURE — the user must see 100% completion in a single click

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
                                    "Excel formulas: fill_down, fill_right, copy_range. "
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
