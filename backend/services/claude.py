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
You are tsifl, an elite AI financial analyst and workflow assistant embedded inside Excel, RStudio, Terminal, and Gmail.
You can READ the user's current environment and make REAL changes in real time across all four apps.
You have shared memory across all apps — everything you know in Excel is available in R, Terminal, and Gmail.

## Your Capabilities
### Excel
- Write individual cells or entire ranges
- Build full financial model structures (LBO, DCF, 3-statement, comps)
- Read existing data and analyze or extend it
- Format cells (bold headers, color coding, number formats)

### RStudio
- Write and execute R code in the user's environment
- Insert code into the active script editor

### Terminal
- Execute shell commands, write files, open URLs

### Gmail
- Read inbox, search emails, draft and send emails

## Financial Model Guidelines
When building model structures in Excel:
- Header rows: color "#0a1929" background, white font_color, bold true
- Number format "$#,##0" for currency, "0.0%" for percentages
- Years across columns (B, C, D...), line items down rows (A column)
- Always call autofit as the final action

## Rules
- Reply in ONE short sentence (max 15 words). No bullet points. No explanations.
- Always call the execute_actions tool — never output raw JSON as text.
- If the sheet already has data, respect its structure.
- Never make up financial data — use 0 or "TBD" as placeholders.
- For R code, write clean, runnable code only.
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
                    "description": "List of actions to execute in order.",
                    "items": {
                        "type": "object",
                        "required": ["type", "payload"],
                        "properties": {
                            "type": {
                                "type": "string",
                                "description": (
                                    "Action type. Excel: write_cell, write_range, format_range, autofit. "
                                    "R: run_r_code, install_package. "
                                    "Terminal: run_shell_command, write_file, open_url. "
                                    "Gmail: send_email, draft_email, search_emails."
                                )
                            },
                            "payload": {
                                "type": "object",
                                "description": (
                                    "Action parameters. "
                                    "write_cell: {cell, value, bold?, color?, font_color?}. "
                                    "write_range: {range, values (2D array), bold?, color?, font_color?}. "
                                    "format_range: {range, bold?, color?, font_color?, number_format?}. "
                                    "autofit: {}. "
                                    "run_r_code: {code}. "
                                    "install_package: {package}. "
                                    "run_shell_command: {command}. "
                                    "write_file: {path, content}. "
                                    "open_url: {url}. "
                                    "send_email: {to, subject, body}. "
                                    "draft_email: {to, subject, body}. "
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
        max_tokens  = 4096,
        system      = SYSTEM_PROMPT,
        tools       = TOOLS,
        tool_choice = {"type": "any"},   # Force Claude to always call execute_actions
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
        lines = ["[EXCEL CONTEXT]"]
        lines.append(f"Sheet: {context.get('sheet', 'Sheet1')}")
        lines.append(f"Selected cell: {context.get('selected_cell', 'A1')}")

        sheet_data = context.get("sheet_data", [])
        if sheet_data:
            lines.append(f"Used range: {context.get('used_range', '')}")
            lines.append("Current data:")
            for r_idx, row in enumerate(sheet_data[:30]):
                for c_idx, val in enumerate(row[:20]):
                    if val not in (None, "", 0):
                        lines.append(f"  {_col_letter(c_idx)}{r_idx+1}: {val}")
        else:
            lines.append("Sheet is empty.")

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
