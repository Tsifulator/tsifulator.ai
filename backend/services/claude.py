"""
Claude Service — the AI brain.
Sends user messages + full spreadsheet context to Claude.
Returns a human reply + one or more structured actions to execute.
"""

import anthropic
import os
import json
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).parent.parent.parent / ".env", override=True)

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

SYSTEM_PROMPT = """
You are Tsifulator, an elite AI financial analyst and workflow assistant embedded inside Excel, RStudio, Terminal, and Gmail.
You can READ the user's current environment and make REAL changes in real time across all four apps.
You have shared memory across all apps — everything you know in Excel is available in R, Terminal, and Gmail.

## Your Capabilities
### Excel
- Write individual cells or entire ranges
- Build full financial model structures (LBO, DCF, 3-statement, comps)
- Read existing data and analyze or extend it
- Format cells (bold headers, color coding, number formats)

### RStudio
- Write and execute R code
- Access user's global environment, loaded packages, working directory
- Insert code into the active script editor

### Terminal
- Execute shell commands
- Write files
- Open URLs

### Gmail
- Read the inbox and search emails
- Draft professional emails
- Send emails

## Response Format
Always respond with:
1. A short, confident reply (1-2 sentences max — you're in a sidebar or terminal)
2. A JSON block with your actions

For a SINGLE action:
```json
{
  "type": "write_cell",
  "payload": {"cell": "B2", "value": 50000}
}
```

For MULTIPLE actions (e.g. building a model):
```json
{
  "actions": [
    {"type": "write_range", "payload": {"range": "A1:D1", "values": [["Revenue", "COGS", "GP", "EBITDA"]], "bold": true, "color": "#0a1929", "font_color": "#ffffff"}},
    {"type": "autofit", "payload": {}}
  ]
}
```

## Excel Action Types
- write_cell: {"cell": "B2", "value": ..., "bold": true/false, "color": "#hex", "font_color": "#hex"}
- write_range: {"range": "A1:D1", "values": [[...]], "bold": true/false, "color": "#hex", "font_color": "#hex"}
- format_range: {"range": "A1:D1", "bold": true/false, "color": "#hex", "number_format": "$#,##0"}
- autofit: {}

## R Action Types
- run_r_code: {"code": "...valid R code..."}
- install_package: {"package": "packagename"}

## Terminal Action Types
- run_shell_command: {"command": "...shell command..."}
- write_file: {"path": "relative/path.txt", "content": "...file contents..."}
- open_url: {"url": "https://..."}

## Gmail Action Types
- send_email: {"to": "...", "subject": "...", "body": "..."}
- draft_email: {"to": "...", "subject": "...", "body": "..."}
- search_emails: {"query": "from:someone subject:topic"}

## Financial Model Guidelines
When building model structures in Excel:
- Header rows: color "#0a1929" background, white text, bold
- Number format "$#,##0" for currency, "0.0%" for percentages
- Years across columns (B, C, D...), line items down rows (A column)
- Always autofit at the end

## Terminal Guidelines
- For shell commands: always use safe, non-destructive commands unless user explicitly asks
- For file writes: use relative paths unless absolute is specified
- Never run rm -rf or destructive commands without explicit user instruction

## Gmail Guidelines
When drafting emails, be professional and concise.
Subject lines should be clear and direct.
Sign off with the user's name if known from context.
For financial deal correspondence, be formal.

## Shared Memory
The conversation history shows which app each message came from.
Use this to connect context across apps — e.g., reference an LBO model built in Excel
when answering in Terminal, or draft an email summarizing a model built in RStudio.

## Rules
- Be concise. One or two sentences max. No essays.
- Always output a JSON block when taking action.
- If the sheet already has data, respect its structure.
- Never make up financial data — use 0 or "TBD" as placeholders.
- For R code, write clean, runnable code only.
- For shell commands, prefer safe, readable one-liners.
"""

async def get_claude_response(message: str, context: dict,
                              session_id: str, history: list = []) -> dict:
    # Format current app context
    sheet_summary = _format_sheet_context(context)
    user_content = f"{message}\n\n{sheet_summary}" if sheet_summary else message

    # Build message thread: history + current message
    # History gives Claude cross-app, cross-session memory
    messages = []
    for h in history:
        role = h.get("role", "user")
        content = h.get("content", "")
        app = h.get("app", "")
        if role == "user" and app:
            content = f"[From {app}] {content}"
        messages.append({"role": role, "content": content})

    # Add current message
    messages.append({"role": "user", "content": user_content})

    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2048,
        system=SYSTEM_PROMPT,
        messages=messages
    )

    raw = response.content[0].text
    return _parse_response(raw)


def _format_sheet_context(context: dict) -> str:
    """Formats Excel or RStudio context into a clean summary for Claude."""
    if not context:
        return ""

    app = context.get("app", "excel")

    if app == "excel":
        lines = ["[EXCEL CONTEXT]"]
        lines.append(f"Sheet: {context.get('sheet', 'Sheet1')}")
        lines.append(f"Selected cell: {context.get('selected_cell', 'A1')}")

        sheet_data = context.get("sheet_data", [])
        if sheet_data:
            lines.append(f"Used range: {context.get('used_range', '')}")
            lines.append("Current data (row, col, value):")
            for r_idx, row in enumerate(sheet_data[:30]):
                for c_idx, val in enumerate(row[:20]):
                    if val not in (None, "", 0):
                        col_letter = _col_letter(c_idx)
                        lines.append(f"  {col_letter}{r_idx + 1}: {val}")
        else:
            lines.append("Sheet is empty.")

    elif app == "rstudio":
        lines = ["[RSTUDIO CONTEXT]"]
        lines.append(f"R version: {context.get('r_version', 'unknown')}")
        lines.append(f"Working dir: {context.get('working_dir', '~')}")
        lines.append(f"Loaded packages: {context.get('loaded_pkgs', 'none')}")

        env_objects = context.get("env_objects", [])
        if env_objects:
            lines.append("Global environment objects:")
            for obj in env_objects:
                dim = f" [{obj.get('dim')}]" if obj.get('dim') else ""
                lines.append(f"  {obj['name']} ({obj['class']}{dim}): {obj.get('preview', '')}")
        else:
            lines.append("Global environment is empty.")

    elif app == "terminal":
        lines = ["[TERMINAL CONTEXT]"]
        lines.append(f"Shell: {context.get('shell', 'bash')}")
        lines.append(f"Working dir: {context.get('working_dir', '~')}")
        lines.append(f"User: {context.get('user', '')}")
        lines.append(f"OS: {context.get('os', 'darwin')}")

        recent = context.get("recent_commands", [])
        if recent:
            lines.append("Recent commands:")
            for cmd in recent:
                lines.append(f"  {cmd}")

        ls_files = context.get("ls", [])
        if ls_files:
            lines.append(f"Current directory files: {', '.join(ls_files[:15])}")

    elif app == "gmail":
        lines = ["[GMAIL CONTEXT]"]
        lines.append(f"Account: {context.get('email', 'connected')}")

        recent_emails = context.get("recent_emails", [])
        if recent_emails:
            lines.append("Recent inbox emails:")
            for email in recent_emails[:5]:
                lines.append(f"  From: {email.get('from', '')} | Subject: {email.get('subject', '')} | {email.get('date', '')}")

        current_email = context.get("current_email")
        if current_email:
            lines.append(f"\nCurrent email open:")
            lines.append(f"  From: {current_email.get('from', '')}")
            lines.append(f"  Subject: {current_email.get('subject', '')}")
            lines.append(f"  Body: {current_email.get('body', '')[:500]}")
    else:
        return ""

    return "\n".join(lines)


def _col_letter(idx: int) -> str:
    """Converts column index (0-based) to Excel letter (A, B, ... Z, AA...)."""
    letters = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _parse_response(raw: str) -> dict:
    """Parses Claude's response into reply text + action(s)."""
    action = {}
    actions = []
    reply = raw

    try:
        if "```json" in raw:
            json_str = raw.split("```json")[1].split("```")[0].strip()
            parsed = json.loads(json_str)
            reply = raw.split("```json")[0].strip()

            # Handle multi-action responses
            if "actions" in parsed:
                actions = parsed["actions"]
            else:
                action = parsed
    except Exception:
        reply = raw

    return {"reply": reply, "action": action, "actions": actions}
