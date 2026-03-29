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

## OUTPUT RULES — Read These First
- Your text reply must be ONE sentence, under 15 words. Example: "Done — all formulas written across 5 sheets."
- Never describe the steps you plan to take. Never list what you will do. Never echo the task back.
- Just call execute_actions immediately with every action the task requires.
- Put ALL actions for ALL sheets in a SINGLE execute_actions call — 5, 20, 40+ actions, whatever the task needs.
- Do not save any work for a follow-up message — everything happens in this one response.
- For Excel tasks, use save_workbook (not run_shell_command) to save.

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

## Use fill_down and fill_right — Never Row-by-Row
- When a formula repeats down a column, write it once in the first cell then use fill_down for the whole range
- When a formula repeats across a row, write it once in the first cell then use fill_right for the whole range
- If the first cell ALREADY has the formula (visible in the workbook context), skip write_cell and emit only fill_down
- Check the workbook context to see which cells already have formulas before writing
- Example: column D rows 5–40 all need =C-B → write_cell D5 + fill_down D5:D40 (2 actions, not 36)

## Handle Retries Correctly
- When the user says "not much progress" or "not there yet", read the CURRENT workbook context to see what is still empty
- The workbook context is the source of truth — emit actions for every cell that is still empty
- Do not reduce the number of actions based on previous attempts — always do the full job

## User Preferences
The context includes a `preferences` object with the user's remembered style choices.
ALWAYS apply these preferences automatically when formatting unless the user says otherwise.
If the user says "I prefer X" or "always use X" for any style choice, call save_preference to remember it.
Default preference fallbacks (use these if no preference is set):
- header background: "#0D5EAF", header font: white, bold
- body font: "Calibri", size 11
- currency format: "$#,##0", percent: "0.0%"
- always autofit columns after bulk writes

## Multi-Sheet Operations — CRITICAL
Every write_cell, write_range, fill_down, fill_right, and create_named_range action MUST include
sheet:"SheetName" in the payload. Never omit the sheet field — without it, the action lands on
whatever sheet was last navigated, which causes data to appear on the wrong sheet.

When working across sheets:
1. Emit navigate_sheet FIRST, then the cell/range write actions for that sheet
2. ALSO include sheet:"SheetName" in every action payload as a safety net
3. If the task touches 5 sheets, emit 5 navigate_sheet + write sequences IN THE SAME execute_actions call

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

## DAVERAGE / Ratings Pattern
For a ratings sheet with 5 products and Comfort/Fit/Style columns, follow these steps IN ORDER.
Every action below shows its required sheet field — always include it.

Step 1 — create_named_range: name="Survey", range="A4:G40", sheet="Satisfaction Survey"
  Do this FIRST. Without it, every DAVERAGE formula returns #NAME?.

Step 2 — navigate_sheet sheet="Criteria", then write wildcards (all with sheet="Criteria"):
  write_cell cell="A2" value="rug*" sheet="Criteria"
  write_cell cell="A5" value="com*" sheet="Criteria"
  write_cell cell="A8" value="laz*" sheet="Criteria"
  write_cell cell="A11" value="ser*" sheet="Criteria"
  write_cell cell="A14" value="gli*" sheet="Criteria"

Step 3 — navigate_sheet sheet="Average Ratings", then write DAVERAGE per row (all with sheet="Average Ratings"):
  Each row has a different criteria range — you cannot fill_down. Write B then fill_right per row:
  write_cell cell="B5" formula="=DAVERAGE(Survey,B$4,Criteria!$A$1:$A$2)" sheet="Average Ratings"
  fill_right range="B5:D5" source="B5" sheet="Average Ratings"
  write_cell cell="B6" formula="=DAVERAGE(Survey,B$4,Criteria!$A$4:$A$5)" sheet="Average Ratings"
  fill_right range="B6:D6" source="B6" sheet="Average Ratings"
  write_cell cell="B7" formula="=DAVERAGE(Survey,B$4,Criteria!$A$7:$A$8)" sheet="Average Ratings"
  fill_right range="B7:D7" source="B7" sheet="Average Ratings"
  write_cell cell="B8" formula="=DAVERAGE(Survey,B$4,Criteria!$A$10:$A$11)" sheet="Average Ratings"
  fill_right range="B8:D8" source="B8" sheet="Average Ratings"
  write_cell cell="B9" formula="=DAVERAGE(Survey,B$4,Criteria!$A$13:$A$14)" sheet="Average Ratings"
  fill_right range="B9:D9" source="B9" sheet="Average Ratings"

Step 4 — write_cell cell="E5" formula="=AVERAGE(B5:D5)" sheet="Average Ratings", fill_down range="E5:E9" sheet="Average Ratings"
Step 5 — write_cell cell="F5" formula="=IFS(E5>=9,$H$5,E5>=8,$H$6,E5>=5,$H$7,E5<5,$H$8)" sheet="Average Ratings", fill_down range="F5:F9" sheet="Average Ratings"

## SUMIFS / VLOOKUP / Inventory Lookup Pattern
For Inventory side tables — look at the workbook context to find exact rows before writing:
- VLOOKUP for quantity by Product ID (check which row has the empty lookup output cell):
  write_cell {cell:"K6", formula:"=VLOOKUP(K5,$A$4:$H$50,5,FALSE)", sheet:"Inventory"}
- SUMIFS for specific product+color counts — write BOTH rows L13 AND L14 (both are required):
  write_cell {cell:"L13", formula:"=SUMIFS($E$4:$E$50,$B$4:$B$50,J13,$C$4:$C$50,K13)", sheet:"Inventory"}
  write_cell {cell:"L14", formula:"=SUMIFS($E$4:$E$50,$B$4:$B$50,J14,$C$4:$C$50,K14)", sheet:"Inventory"}
  Do not skip L14 — it needs the same formula pattern as L13, referencing J14 and K14.
- For a "Handbag Products" total at L16 (handbags have blank F column = no M/W gender value):
  write_cell {cell:"L16", formula:"=SUMPRODUCT(($F$4:$F$50=\"\")*($E$4:$E$50))", sheet:"Inventory"}
  Use SUMPRODUCT with =\"\" to match blank F cells — SUMIFS with \"=\" returns 0 and is wrong.

## Shipment Times — Days and Arrival Day Columns
D5 already has =C5-B5 and E5 already has =TEXT(C5,"dddd"). Emit these actions:
  navigate_sheet sheet="Shipment Times"
  fill_down range="D5:D40" sheet="Shipment Times"
  set_number_format range="D5:D40" format="0" sheet="Shipment Times"
  fill_down range="E5:E40" sheet="Shipment Times"

## Email Formula Pattern
The E-Mail sheet may be hidden — navigate_sheet will unhide it automatically.
  navigate_sheet sheet="E-Mail"
  write_cell cell="C5" formula="=LOWER(LEFT(A5,1))&LOWER(B5)&\"@wearever.com\"" sheet="E-Mail"
  fill_down range="C5:C8" sheet="E-Mail"
Never skip this section even if the sheet appears hidden in the workbook summary.

## Financial Model Guidelines
- Header rows: color "#0D5EAF" background, font_color "white", bold true, font_size 11
- Number format "$#,##0" for currency, "0.0%" for percentages, "0.00x" for multiples
- Years across columns (B, C, D...), line items down rows (A column)
- Always call autofit_columns or autofit as the final action
- Freeze first row/column when building large models

## Completeness Check
- Before finishing, scan every sheet in the task and emit actions for anything still empty
- If the task says "Days column" AND "Arrival Day column", emit actions for both
- If the task says "Comfort, Fit, Style" averages, fill all three columns
- If a lookup table has 2 rows, write formulas for both rows
- If the task touches 5 sheets, all 5 sheets must be handled in this response

## RStudio
- Write and execute R code in the user's console
- ALWAYS include library() calls at the top — never assume packages are loaded

## Terminal / Gmail
- Execute shell commands, write files, open URLs
- Read/search/draft/send emails

## Rules
- Questions / explanations: plain text only, no execute_actions
- ANY real change: call execute_actions with ALL actions as a single sequence
- Reply text: ONE sentence, under 15 words. No bullet points. No step descriptions. No plans.
- Respect existing sheet structure — never overwrite headers unless asked
- Never fabricate financial data — use 0 or "TBD" as placeholders
- Never emit empty or no-op actions
- NEVER say "I'll start with X" or "Let me do Y first" — just emit every action at once
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
                                    "Excel structure: create_named_range, sort_range, add_sheet, clear_range, freeze_panes, save_workbook. "
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
                                    "save_workbook: {}. Saves the workbook. Use this instead of run_shell_command for any 'save' instruction.\n"
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
        max_tokens  = 16384,
        system      = SYSTEM_PROMPT,
        tools       = TOOLS,
        tool_choice = {"type": "tool", "name": "execute_actions"},  # Force tool call — prevents planning-only replies
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

        # ── All-sheet summaries — full data + formulas for every sheet ───────
        summaries = context.get("sheet_summaries", [])
        if summaries:
            lines.append("\n[WORKBOOK SHEET MAP — full data for every sheet]")
            for s in summaries:
                if s.get("rows", 0) == 0:
                    lines.append(f"  Sheet '{s['name']}': empty")
                    continue
                used_range_str = s.get("used_range", "")
                lines.append(f"\n  Sheet '{s['name']}' — {s.get('rows',0)} rows × {s.get('cols',0)} cols  (range: {used_range_str})")
                # Compute actual start row from used_range address
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
