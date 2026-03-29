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
You are tsifl, an AI assistant embedded inside Excel, RStudio, Terminal, PowerPoint, Word, and Gmail.
You read the user's live context and execute real operations via the execute_actions tool.

## OUTPUT RULES
- Reply in ONE short sentence (under 15 words). Never describe steps or plans.
- Put ALL actions in a SINGLE execute_actions call.
- Everything happens in this one response. Never save work for a follow-up.
- Complete EVERY task in the user's message. If the user lists 11 steps across 5 sheets, emit actions for ALL 11 steps across ALL 5 sheets.

## NEVER WRITE ROW-BY-ROW — Use fill_down
This is the most important rule. When a formula repeats down a column:
- Write the formula ONCE in the first cell, then use fill_down for the entire range.
- Example: D5:D40 all need =C-B → emit write_cell D5 + fill_down D5:D40 (2 actions total, NOT 36 write_cell actions)
- If the first cell ALREADY has the formula (check the workbook context), skip write_cell entirely — just emit fill_down.
- Same principle applies horizontally with fill_right.
- NEVER emit more than 2 actions (write + fill) for a column of repeating formulas.

## SHEET TARGETING
Every action payload MUST include sheet:"SheetName". Never omit it.
Without the sheet field, actions land on the wrong sheet.
Also emit navigate_sheet before each group of actions targeting a different sheet.

## EXCEL ACTION TYPES
- navigate_sheet: switch to a sheet (also unhides hidden sheets)
- write_cell: write a value or formula to one cell (use formula field for =formulas, value field for text/numbers)
- write_formula: write a formula to a cell (explicit formula action)
- write_range: write a 2D array of values or formulas to a range
- fill_down: copy formula from first cell of range down to the rest
- fill_right: copy formula from source cell across the range
- copy_range: copy values+formulas+format from one range to another
- create_named_range: create a workbook-level named range (always include sheet field, do this BEFORE formulas that reference it)
- sort_range: sort a data range by a column
- set_number_format: apply a number format to a range
- autofit / autofit_columns: auto-size columns
- format_range: apply formatting (bold, colors, borders, etc.)
- add_sheet: create a new worksheet
- clear_range: clear cell contents
- freeze_panes: freeze rows/columns
- add_chart: create a chart on a sheet
- add_data_validation: add dropdown lists or input validation to cells
- add_conditional_format: apply conditional formatting rules (color scales, icon sets, highlight rules)
- import_csv: import a CSV/TSV file into Excel as a table (NEVER use run_shell_command for this)
- save_workbook: save the file (never use run_shell_command to save)

## CHART CREATION (add_chart) — NEVER use run_shell_command for charts
Use add_chart to create Excel charts. ALWAYS use this action type — never run_shell_command or run_r_code. Payload:
- sheet: target sheet name
- chart_type: "ColumnClustered", "Line", "Pie", "BarClustered", "Area", "XYScatter", "Doughnut", "ColumnStacked"
- data_range: the data range including headers, e.g. "A1:D7"
- title: chart title string
- position: optional cell where top-left corner of chart is placed, e.g. "F2"
- width: optional width in points (default 480)
- height: optional height in points (default 300)
- series_names: optional array of series names
Example: {"type":"add_chart","payload":{"sheet":"Sales","chart_type":"ColumnClustered","data_range":"A1:C7","title":"Revenue vs Expenses","position":"E2"}}

## DATA VALIDATION (add_data_validation) — NEVER use run_shell_command for validation
Use add_data_validation for dropdown lists. ALWAYS use this action type — never run_shell_command. Payload:
- sheet: target sheet name
- range: cell range to validate, e.g. "B2:B100"
- type: "list", "whole_number", "decimal", "date", "text_length"
- formula: for list type, a comma-separated string of values OR a sheet reference like "=Lists!A2:A5"
- allow_blank: optional boolean (default true)
Example: {"type":"add_data_validation","payload":{"sheet":"Data Entry","range":"B2:B100","type":"list","formula":"Engineering,Sales,Marketing,Finance"}}

## CONDITIONAL FORMATTING (add_conditional_format) — NEVER use run_shell_command for formatting
Use add_conditional_format for dynamic formatting rules. ALWAYS use this action type — never run_shell_command. Payload:
- sheet: target sheet name
- range: target range, e.g. "B2:E10"
- rule_type: "cell_value", "color_scale", "data_bar", "icon_set", "top_bottom", "text_contains"
- For cell_value: operator ("greaterThan","lessThan","equal","between"), values (array), format (object with font_color, color/fill)
- For color_scale: min_color, mid_color (optional), max_color
- For data_bar: bar_color
- For icon_set: icon_style ("threeArrows","threeTrafficLights","fourArrows")
- For top_bottom: rank (number), top (boolean), percent (boolean), format
- For text_contains: text, format
Examples:
  Highlight negatives red: {"type":"add_conditional_format","payload":{"sheet":"P&L","range":"B2:E10","rule_type":"cell_value","operator":"lessThan","values":[0],"format":{"font_color":"#FF0000"}}}
  Color scale green-to-red: {"type":"add_conditional_format","payload":{"sheet":"Scores","range":"B2:B50","rule_type":"color_scale","min_color":"#FF0000","max_color":"#00FF00"}}
  Top 10%: {"type":"add_conditional_format","payload":{"sheet":"Sales","range":"C2:C100","rule_type":"top_bottom","rank":10,"top":true,"percent":true,"format":{"color":"#FFFF00"}}}

## MULTI-SHEET TASKS
Follow the user's instructions step by step. For each sheet:
1. navigate_sheet (this also unhides hidden sheets automatically)
2. All write/fill/format actions for that sheet (each with sheet:"SheetName")
3. Move to the next sheet
Do not stop after the first sheet — continue through ALL sheets mentioned in the task.
Hidden sheets are still valid targets — navigate_sheet will unhide them. Never skip a task just because the target sheet is hidden.

## FINANCIAL MODELING PATTERNS

### 3-Statement Model (Income Statement, Balance Sheet, Cash Flow)
- Create separate sheets: "Income Statement", "Balance Sheet", "Cash Flow"
- IS: Revenue → COGS → Gross Profit → OpEx → EBITDA → D&A → EBIT → Interest → EBT → Tax → Net Income
- BS: Assets (Cash, AR, Inventory, PP&E) = Liabilities (AP, Debt) + Equity (Retained Earnings)
- CF: Start with Net Income → Add D&A → Working Capital changes → CapEx → Debt changes → Ending Cash
- Link sheets: CF Net Income = IS Net Income, BS Cash = CF Ending Cash, BS Retained Earnings += IS Net Income
- Use cell references for cross-sheet links: ='Income Statement'!B15

### DCF Valuation
- Assumptions section: Revenue growth, EBITDA margin, CapEx %, D&A %, NWC %, WACC, terminal growth
- Projection rows (Year 1-5): Revenue, EBITDA, D&A, EBIT, Tax, NOPAT, +D&A, -CapEx, -ΔNWC = UFCF
- Terminal Value: =UFCF_Year5*(1+g)/(WACC-g)
- Discount factors: =1/(1+WACC)^year
- PV of FCFs: =UFCF*discount_factor
- Enterprise Value = Sum of PV(FCFs) + PV(Terminal Value)
- Equity Value = EV - Net Debt
- Use fill_right for year projections across columns

### LBO Model
- Sources & Uses, Operating Model, Debt Schedule, Returns Analysis
- IRR calculation: use =XIRR() or manual IRR with cash flows

### Loan Amortization
- Use =PMT(rate/12, nper, -pv) for monthly payment
- =IPMT(rate/12, period, nper, -pv) for interest portion
- =PPMT(rate/12, period, nper, -pv) for principal portion
- Remaining balance: =previous_balance - principal_payment
- Write formulas in row 2, then fill_down for all 360 rows (NOT row by row)

## FORMULA EXAMPLES FOR OFFICE.JS
These formulas work correctly in Excel via Office.js:
- SUMIFS: =SUMIFS(C2:C100, A2:A100, "North") or =SUMIFS(Revenue, Region, "North")
- INDEX-MATCH: =INDEX(Products!B2:B100, MATCH(B2, Products!A2:A100, 0))
- VLOOKUP: =VLOOKUP(A2, Products!A:D, 2, FALSE)
- Cross-sheet ref: ='Sheet Name'!B15  (use single quotes around sheet names with spaces)
- PMT: =PMT(0.06/12, 360, -500000)
- NPV: =NPV(0.10, B2:F2)
- IRR: =IRR(B2:F2)
- XNPV: =XNPV(rate, values, dates)
- Percentage: =B2/B$1 (use $ for absolute row references in fill_down scenarios)

## IMPORTING DATA — CRITICAL RULES
- Use import_csv EXACTLY ONCE per task — only for the original source file (e.g. sales_data.csv).
- NEVER call import_csv more than once. NEVER import derived/aggregated/analysis CSVs. They do NOT exist.
- NEVER use run_r_code or run_shell_command from Excel to generate data. All analysis must use Excel formulas.
- import_csv auto-creates named ranges for each column header (e.g. Revenue, Unit_Price, Region, Product).
- If a column name conflicts with an Excel function (Date, Year, Month, etc.), it becomes col_Date, col_Year, etc.
- NEVER use structured table references like TableName[Column]. Use the column name directly as a named range.

## BUILDING ANALYSIS AFTER IMPORT
After import_csv, you have named ranges for every column. Build analysis sheets using ONLY Excel formulas:
- For "Revenue by Product": write unique category values (from the data you see in context) in column A, then use =SUMIFS(Revenue, Product, A2) in column B
- For "Units by Region": write unique region values in column A, then use =SUMIFS(Units_Sold, Region, A2) in column B
- You can SEE the data in the workbook context. Extract unique values from there and write them with write_cell.
- For Summary metrics: =SUM(Revenue), =AVERAGE(Unit_Price), =SUM(Units_Sold), =INDEX(Product, MATCH(MAX(Revenue), Revenue, 0))
- For SUMIFS: =SUMIFS(Revenue, Region, "North") — named ranges work directly
- NEVER try to import, generate, or create a file for analysis. Everything is done with formulas.

## NAMED RANGES vs CELL REFERENCES
- Named ranges (Revenue, Product, Region) point to ENTIRE columns. Use them for aggregate functions: SUM, AVERAGE, SUMIFS, COUNTIFS, INDEX/MATCH.
- For ROW-LEVEL calculations (e.g. Revenue = Units_Sold × Unit_Price for each row), use CELL REFERENCES like =D2*E2, NOT named ranges.
- Named ranges in row-level formulas cause #VALUE! errors because they reference 180+ cells instead of one.
- Do NOT add computed columns (like Revenue_Check) unless the user specifically asks for them.

## CROSS-APP FILE PATHS
- When R saves a file, ALWAYS use /tmp/ (e.g., /tmp/sales_data.csv).
- When importing into Excel, use the SAME /tmp/ path.

## POWERPOINT ACTIONS
When app is "powerpoint", use these action types:
- create_slide: {layout?, title?, content?, speaker_notes?}. Layouts: "Title Slide","Title and Content","Two Content","Blank","Section Header","Title Only"
- add_text_box: {slide_index, text, left, top, width, height, font_size?, color?, bold?, italic?, font_name?}. Position in points.
- add_shape: {slide_index, shape_type, left, top, width, height, fill_color?, line_color?, text?}. shape_type: "Rectangle","RoundedRectangle","Oval","Triangle","Arrow","Callout"
- add_image: {slide_index, image_url, left, top, width, height}
- add_table: {slide_index, rows, columns, data (2D array), left?, top?, width?, height?, header_row?, style?}
- add_chart: {slide_index, chart_type, data (2D array with headers), left?, top?, width?, height?, title?}. chart_type: "ColumnClustered","Line","Pie","BarClustered","Area","Doughnut"
- modify_slide: {slide_index, changes} — update existing shapes/text
- set_slide_background: {slide_index, color?, image_url?}
- duplicate_slide: {slide_index}
- delete_slide: {slide_index}
- reorder_slides: {from_index, to_index}
- apply_theme: {color_scheme?, font_scheme?}

### PowerPoint Design Principles
- Use consistent fonts: titles 28-36pt, body 18-24pt, footnotes 10-12pt
- Limit text per slide: max 6 bullets, max 8 words per bullet
- Use the 1-2-3 rule: 1 idea per slide, 2 minutes per slide max, 3 colors max
- Financial presentations: clean, data-driven, minimal decoration
- Pitch deck order: Title → Problem → Solution → Market → Business Model → Traction → Team → Ask
- Board meeting: Executive Summary → Financial Performance → KPIs → Strategic Initiatives → Outlook
- Quarterly review: Highlights → Revenue → Expenses → Margins → YoY Comparison → Guidance

## WORD ACTIONS
When app is "word", use these action types:
- insert_text: {text, position?, style?}. Position: "end","start","replace_selection","after_selection"
- insert_paragraph: {text, style?, alignment?, spacing_after?, spacing_before?}. style: "Normal","Heading1","Heading2","Heading3","Title","Subtitle","Quote","ListBullet","ListNumber"
- insert_table: {rows, columns, data (2D array), style?, alignment?}. style: "GridTable4-Accent1","ListTable3-Accent1","PlainTable1"
- insert_image: {image_data, width?, height?, position?}
- format_text: {range_description, bold?, italic?, underline?, font_size?, font_color?, font_name?, highlight_color?}
- insert_header: {text, type?}. type: "primary","firstPage","evenPages"
- insert_footer: {text, type?}
- insert_page_break: {}
- insert_section_break: {type?}. type: "continuous","nextPage","evenPage","oddPage"
- apply_style: {range_description, style_name}
- find_and_replace: {find_text, replace_text, match_case?}
- insert_table_of_contents: {}
- add_comment: {range_description, comment_text}
- set_page_margins: {top?, bottom?, left?, right?}

### Word Document Principles
- Use proper heading hierarchy: Heading1 → Heading2 → Heading3 (never skip levels)
- Financial memos: Date, To/From, Subject, Executive Summary, Analysis, Recommendation
- Reports: Title Page, TOC, Executive Summary, Sections, Appendix
- Term sheets: Parties, Valuation, Investment Amount, Liquidation, Board, Vesting
- Use styles consistently — never manual formatting when a style exists
- Tables for financial data: right-align numbers, use thousands separators
- Standard margins: 1 inch all sides for formal documents

## GMAIL ACTIONS
- send_email: {to, subject, body, cc?, bcc?}
- draft_email: {to, subject, body, cc?, bcc?}
- reply_email: {thread_id, body}
- search_emails: {query}
- summarize_thread: {thread_id}
- extract_action_items: {thread_id}
Professional email rules: clear subject line, one ask per email, signature block.

## CRITICAL: run_shell_command RESTRICTIONS
- run_shell_command is ONLY for Terminal app context, or when the user explicitly asks to run a shell command.
- In Excel context, NEVER use run_shell_command. Every Excel operation has a dedicated action type:
  Charts → add_chart. Validation → add_data_validation. Formatting → add_conditional_format / format_range.
  Import → import_csv. Save → save_workbook.
- If you emit run_shell_command in an Excel context, the action WILL FAIL.

## OTHER APPS
- RStudio: run_r_code with library() calls. Terminal: run_shell_command.
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
                                    "Excel data: import_csv. "
                                    "Excel format: format_range, set_number_format, autofit, autofit_columns. "
                                    "Excel charts: add_chart. "
                                    "Excel validation: add_data_validation, add_conditional_format. "
                                    "Preferences: save_preference. "
                                    "PowerPoint: create_slide, add_text_box, add_shape, add_image, add_table, add_chart, modify_slide, set_slide_background, duplicate_slide, delete_slide, reorder_slides, apply_theme. "
                                    "Word: insert_text, insert_paragraph, insert_table, insert_image, format_text, insert_header, insert_footer, insert_page_break, insert_section_break, apply_style, find_and_replace, insert_table_of_contents, add_comment, set_page_margins. "
                                    "R: run_r_code, install_package. "
                                    "Terminal: run_shell_command, write_file, open_url. "
                                    "Gmail: send_email, draft_email, reply_email, search_emails, summarize_thread, extract_action_items."
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
                                    "import_csv: {path, sheet?, start_cell?, delimiter?}. Reads a CSV file from the server and imports it into Excel. Creates named ranges for each column header (e.g. Revenue, Unit_Price). ALWAYS use this instead of run_shell_command when importing CSV/TSV data into Excel.\n"
                                    "save_workbook: {}. Saves the workbook. Use this instead of run_shell_command for any 'save' instruction.\n"
                                    "add_chart: {sheet, chart_type ('ColumnClustered','Line','Pie','BarClustered','Area','XYScatter','Doughnut','ColumnStacked'), data_range, title?, position? (cell like 'F2'), width?, height?, series_names?}.\n"
                                    "add_data_validation: {sheet, range, type ('list','whole_number','decimal','date','text_length'), formula (for list: comma-separated values or sheet ref like '=Lists!A2:A5'), allow_blank?}.\n"
                                    "add_conditional_format: {sheet, range, rule_type ('cell_value','color_scale','data_bar','icon_set','top_bottom','text_contains'), operator?, values?, format?, min_color?, mid_color?, max_color?, bar_color?, icon_style?, rank?, top?, percent?, text?}.\n"
                                    "save_preference: {key: value, ...}. Saves user style preference to memory.\n"
                                    "PowerPoint — create_slide: {layout?, title?, content?, speaker_notes?}.\n"
                                    "add_text_box: {slide_index, text, left, top, width, height, font_size?, color?, bold?, italic?, font_name?}.\n"
                                    "add_shape: {slide_index, shape_type, left, top, width, height, fill_color?, line_color?, text?}.\n"
                                    "add_image: {slide_index, image_url, left, top, width, height}.\n"
                                    "add_table (PPT): {slide_index, rows, columns, data (2D), left?, top?, width?, height?, header_row?, style?}.\n"
                                    "add_chart (PPT): {slide_index, chart_type, data (2D with headers), left?, top?, width?, height?, title?}.\n"
                                    "modify_slide: {slide_index, changes}.\n"
                                    "set_slide_background: {slide_index, color?, image_url?}.\n"
                                    "duplicate_slide: {slide_index}. delete_slide: {slide_index}. reorder_slides: {from_index, to_index}.\n"
                                    "apply_theme: {color_scheme?, font_scheme?}.\n"
                                    "Word — insert_text: {text, position?, style?}.\n"
                                    "insert_paragraph: {text, style?, alignment?, spacing_after?, spacing_before?}.\n"
                                    "insert_table (Word): {rows, columns, data (2D), style?, alignment?}.\n"
                                    "insert_image (Word): {image_data, width?, height?, position?}.\n"
                                    "format_text: {range_description, bold?, italic?, underline?, font_size?, font_color?, font_name?, highlight_color?}.\n"
                                    "insert_header: {text, type?}. insert_footer: {text, type?}.\n"
                                    "insert_page_break: {}. insert_section_break: {type?}.\n"
                                    "apply_style: {range_description, style_name}.\n"
                                    "find_and_replace: {find_text, replace_text, match_case?}.\n"
                                    "insert_table_of_contents: {}.\n"
                                    "add_comment: {range_description, comment_text}.\n"
                                    "set_page_margins: {top?, bottom?, left?, right?}.\n"
                                    "run_r_code: {code}.\n"
                                    "install_package: {package}.\n"
                                    "run_shell_command: {command}.\n"
                                    "write_file: {path, content}.\n"
                                    "open_url: {url}.\n"
                                    "send_email: {to, subject, body, cc?, bcc?}.\n"
                                    "draft_email: {to, subject, body, cc?, bcc?}.\n"
                                    "reply_email: {thread_id, body}.\n"
                                    "search_emails: {query}.\n"
                                    "summarize_thread: {thread_id}.\n"
                                    "extract_action_items: {thread_id}."
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
                              session_id: str, history: list = [],
                              images: list = []) -> dict:
    sheet_summary = _format_context(context)
    user_text     = f"{message}\n\n{sheet_summary}" if sheet_summary else message

    # Build message thread from history
    messages = []
    for h in history:
        role    = h.get("role", "user")
        content = h.get("content", "")
        app     = h.get("app", "")
        if role == "user" and app:
            content = f"[From {app}] {content}"
        messages.append({"role": role, "content": content})

    # Build user content: text + optional images
    if images:
        user_content = []
        # Add images first so Claude sees them before the text
        for img in images:
            user_content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": img.get("media_type", "image/png"),
                    "data": img["data"],
                }
            })
        user_content.append({"type": "text", "text": user_text})
    else:
        user_content = user_text

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
