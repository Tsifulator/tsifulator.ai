"""
Agent v2 — native Anthropic tool calling with extended thinking and
threaded conversation history.

This is the proper "Claude Code-style" architecture:
  - One tool per action type with strict typed schemas
  - Extended thinking enabled (model reasons before each tool call)
  - tool_use + tool_result blocks threaded into a growing conversation
  - Multi-turn: server keeps state, client executes tools and posts back

The legacy services/claude.py still serves the Excel add-in and other
clients. This module is mounted at /agent/ for the desktop agent.
"""

import os
import json
import time
import logging
from typing import Optional
from pathlib import Path

import anthropic
import httpx
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent.parent / ".env", override=True)

logger = logging.getLogger(__name__)

client = anthropic.Anthropic(
    api_key=os.getenv("ANTHROPIC_API_KEY"),
    timeout=httpx.Timeout(600.0, connect=10.0),
)

MODEL_FAST = "claude-haiku-4-5-20251001"
MODEL_STANDARD = "claude-sonnet-4-20250514"
MODEL_HEAVY = "claude-opus-4-20250514"


# ── Tool schemas — ONE TOOL PER ACTION TYPE ─────────────────────────────────
# Strict input_schema means the API rejects malformed calls before they reach
# the client. No more hallucinated fields.

AGENT_TOOLS = [
    # ── File system ────────────────────────────────────────────────────────
    {
        "name": "search_files",
        "description": (
            "Search the user's Mac via Spotlight, sorted by most recently "
            "modified. Use this for ANY file search — never use shell `find`, "
            "`mdfind`, `ls`, or `locate`."
        ),
        "input_schema": {
            "type": "object",
            "required": ["query"],
            "properties": {
                "query": {"type": "string", "description": "Search term (filename or content)."},
                "file_type": {
                    "type": "string",
                    "enum": ["excel", "word", "ppt", "pdf", "image", "csv", "text"],
                    "description": "Optional filter by file type.",
                },
                "max_results": {"type": "integer", "default": 10},
            },
        },
    },
    {
        "name": "open_file",
        "description": (
            "Open a file in its default app. If the user named a SPECIFIC app "
            "(e.g. 'open this CSV in Excel'), use applescript instead — "
            "open_file uses system defaults (CSV → Numbers, etc)."
        ),
        "input_schema": {
            "type": "object",
            "required": ["path"],
            "properties": {
                "path": {"type": "string", "description": "Absolute file path."},
            },
        },
    },
    {
        "name": "read_file",
        "description": (
            "Read a text file (CSV, TSV, TXT, JSON, .py, .r) from disk. Returns "
            "the contents so you can act on them in the next turn. ONLY use "
            "with absolute paths the user actually provided. NEVER invent a "
            "filename. If the user attached a file (image/PDF), it's already "
            "in your context — do NOT call read_file for it."
        ),
        "input_schema": {
            "type": "object",
            "required": ["path"],
            "properties": {
                "path": {"type": "string", "description": "Absolute path. Must start with /."},
                "max_chars": {"type": "integer", "default": 30000},
            },
        },
    },
    {
        "name": "write_file",
        "description": "Write text content to a file. Creates parent dirs.",
        "input_schema": {
            "type": "object",
            "required": ["path", "content"],
            "properties": {
                "path": {"type": "string"},
                "content": {"type": "string"},
            },
        },
    },

    # ── App control ────────────────────────────────────────────────────────
    {
        "name": "open_app",
        "description": "Launch or activate a Mac app by name.",
        "input_schema": {
            "type": "object",
            "required": ["name"],
            "properties": {
                "name": {"type": "string", "description": "App name (e.g. 'Microsoft Excel', 'Safari')."},
            },
        },
    },
    {
        "name": "open_url",
        "description": "Open a URL in the default browser.",
        "input_schema": {
            "type": "object",
            "required": ["url"],
            "properties": {"url": {"type": "string"}},
        },
    },
    {
        "name": "applescript",
        "description": (
            "Run AppleScript to control any Mac app. Use this for: writing to "
            "Excel/Word/PPT cells, building documents, opening files in a "
            "specific app, pressing Macabacus/Think-Cell shortcuts via "
            "keystroke, etc. Prefer ONE atomic script over multiple actions."
        ),
        "input_schema": {
            "type": "object",
            "required": ["script"],
            "properties": {"script": {"type": "string", "description": "AppleScript source."}},
        },
    },
    {
        "name": "shell",
        "description": (
            "Run a read-only shell command. Do NOT use for file search "
            "(use search_files), file reading (use read_file), or anything "
            "destructive."
        ),
        "input_schema": {
            "type": "object",
            "required": ["command"],
            "properties": {"command": {"type": "string"}},
        },
    },

    # ── Clipboard + notify ─────────────────────────────────────────────────
    {
        "name": "clipboard_copy",
        "description": "Copy text to the system clipboard.",
        "input_schema": {
            "type": "object",
            "required": ["text"],
            "properties": {"text": {"type": "string"}},
        },
    },
    {
        "name": "clipboard_read",
        "description": "Read the current clipboard contents.",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "notify",
        "description": "Show a macOS notification.",
        "input_schema": {
            "type": "object",
            "required": ["message"],
            "properties": {"message": {"type": "string"}},
        },
    },

    # ── Deterministic media/web ────────────────────────────────────────────
    {
        "name": "play_media",
        "description": (
            "Play music/video on a specific platform. Battle-tested — "
            "no vision loop needed."
        ),
        "input_schema": {
            "type": "object",
            "required": ["platform", "query"],
            "properties": {
                "platform": {"type": "string", "enum": ["spotify", "youtube", "apple music"]},
                "query": {"type": "string"},
            },
        },
    },
    {
        "name": "web_search",
        "description": "Open a web search in the browser (just navigation, doesn't return results).",
        "input_schema": {
            "type": "object",
            "required": ["query"],
            "properties": {
                "query": {"type": "string"},
                "engine": {"type": "string", "enum": ["google", "bing", "duckduckgo", "youtube"], "default": "google"},
            },
        },
    },
    {
        "name": "fetch_url",
        "description": "Fetch a URL and return its text content (HTML stripped).",
        "input_schema": {
            "type": "object",
            "required": ["url"],
            "properties": {
                "url": {"type": "string"},
                "max_chars": {"type": "integer", "default": 8000},
            },
        },
    },
    {
        "name": "web_lookup",
        "description": (
            "Search the web AND fetch the top results' content in one call. "
            "Use this when you don't know something (e.g. 'how do I X in "
            "Macabacus') — you'll get actual page text back, not just a URL."
        ),
        "input_schema": {
            "type": "object",
            "required": ["query"],
            "properties": {
                "query": {"type": "string"},
                "max_chars": {"type": "integer", "default": 6000},
            },
        },
    },

    # ── Data export ────────────────────────────────────────────────────────
    {
        "name": "data_export",
        "description": (
            "Export data from a Mac app to a file. Battle-tested per-app "
            "scripts. ALWAYS use this for 'export the data' instead of "
            "writing your own AppleScript."
        ),
        "input_schema": {
            "type": "object",
            "required": ["source_app", "destination"],
            "properties": {
                "source_app": {"type": "string", "description": "e.g. 'Numbers', 'Microsoft Excel'"},
                "destination": {"type": "string", "description": "Output file path."},
                "format": {"type": "string", "enum": ["csv", "tsv", "pdf"], "default": "csv"},
            },
        },
    },

    # ── Gmail ──────────────────────────────────────────────────────────────
    {
        "name": "check_inbox",
        "description": "Fetch recent emails from the user's Gmail inbox.",
        "input_schema": {
            "type": "object",
            "properties": {"max_results": {"type": "integer", "default": 10}},
        },
    },
    {
        "name": "search_email",
        "description": "Search Gmail with query syntax (from:, subject:, is:unread, newer_than:, has:attachment).",
        "input_schema": {
            "type": "object",
            "required": ["query"],
            "properties": {
                "query": {"type": "string"},
                "max_results": {"type": "integer", "default": 10},
            },
        },
    },
    {
        "name": "read_email",
        "description": "Read the full body of a specific email by message ID.",
        "input_schema": {
            "type": "object",
            "required": ["message_id"],
            "properties": {"message_id": {"type": "string"}},
        },
    },
    {
        "name": "send_email",
        "description": (
            "Send an email via Gmail. RED RISK — always requires explicit "
            "user confirmation before being called."
        ),
        "input_schema": {
            "type": "object",
            "required": ["to", "subject", "body"],
            "properties": {
                "to": {"type": "string"},
                "subject": {"type": "string"},
                "body": {"type": "string"},
                "reply_to_id": {"type": "string", "description": "Message ID to thread the reply to."},
            },
        },
    },
    {
        "name": "draft_email",
        "description": "Create a Gmail draft (does NOT send).",
        "input_schema": {
            "type": "object",
            "required": ["to", "subject", "body"],
            "properties": {
                "to": {"type": "string"},
                "subject": {"type": "string"},
                "body": {"type": "string"},
            },
        },
    },

    # ── Screen automation (vision loop) ───────────────────────────────────
    {
        "name": "screenshot",
        "description": "Capture the screen. Returns a base64 image you can analyze in the next turn.",
        "input_schema": {
            "type": "object",
            "properties": {
                "region": {
                    "type": "object",
                    "properties": {
                        "x": {"type": "integer"},
                        "y": {"type": "integer"},
                        "width": {"type": "integer"},
                        "height": {"type": "integer"},
                    },
                },
            },
        },
    },
    {
        "name": "click_at",
        "description": "Click at screen coordinates (from a screenshot).",
        "input_schema": {
            "type": "object",
            "required": ["x", "y"],
            "properties": {
                "x": {"type": "integer"},
                "y": {"type": "integer"},
                "click_type": {"type": "string", "enum": ["left", "right", "double"], "default": "left"},
            },
        },
    },
    {
        "name": "type_text",
        "description": "Type text via keyboard into the focused field.",
        "input_schema": {
            "type": "object",
            "required": ["text"],
            "properties": {"text": {"type": "string"}},
        },
    },
    {
        "name": "key_combo",
        "description": "Press a keyboard shortcut (e.g. 'cmd+c', 'cmd+shift+t', 'return', 'tab', 'escape').",
        "input_schema": {
            "type": "object",
            "required": ["keys"],
            "properties": {"keys": {"type": "string"}},
        },
    },
    {
        "name": "scroll",
        "description": "Scroll the screen.",
        "input_schema": {
            "type": "object",
            "required": ["direction"],
            "properties": {
                "direction": {"type": "string", "enum": ["up", "down", "left", "right"]},
                "amount": {"type": "integer", "default": 3},
            },
        },
    },
    {
        "name": "wait",
        "description": "Wait for UI to settle. Max 10 seconds.",
        "input_schema": {
            "type": "object",
            "required": ["seconds"],
            "properties": {"seconds": {"type": "number", "maximum": 10}},
        },
    },

    # ── Memory & shortcuts ─────────────────────────────────────────────────
    {
        "name": "save_memory",
        "description": (
            "Save a fact about the user that should persist across sessions. "
            "Call this proactively when you learn something useful (email, "
            "preferences, names, project paths)."
        ),
        "input_schema": {
            "type": "object",
            "required": ["fact"],
            "properties": {"fact": {"type": "string"}},
        },
    },
    {
        "name": "set_shortcut",
        "description": (
            "Create a custom shortcut. trigger = slash name; hotkey = "
            "optional system keyboard shortcut (e.g. 'cmd+d')."
        ),
        "input_schema": {
            "type": "object",
            "required": ["trigger", "action"],
            "properties": {
                "trigger": {"type": "string"},
                "action": {"type": "string"},
                "hotkey": {"type": "string"},
            },
        },
    },
]


# ── System prompt (compact — strict rules + thinking guidance) ─────────────
SYSTEM_PROMPT = """You are tsifl, an AI assistant that controls the user's Mac via tools.

## CORE BEHAVIOR

You operate as an iterative agent:
1. The user sends a message (sometimes with attached files or context).
2. You THINK about what they want, what's clear, and what's ambiguous.
3. You either ASK a clarifying question (no tool calls) OR call ONE OR MORE tools.
4. You receive tool_result blocks back, then think again and decide the next step.
5. When the task is complete, respond with a short summary — no more tool calls.

## ASK WHEN AMBIGUOUS — NEVER GUESS

Before calling any tool, check: is the target/file/recipient CRYSTAL CLEAR?
- User attached a file → use that, don't read_file something else
- User typed an absolute path → use exactly that
- User said "the dataset" / "my file" with no path, no attachment, no obvious context → ASK
- User said "in Excel" but didn't say what to put → ASK
- User said "send the email" but no recipient → ASK

Asking is a TEXT response with NO tool calls. The agent loop will give them back to you with their answer.

## NEVER FAKE COMPLETION

Your final text response must match what your tools actually did. Forbidden:
- Saying "I've imported the data" when you only opened the app
- Saying "All set" when you only read a file
- Saying "Done" when no write actions ran

Right pattern: say what's true. "Opened Excel — now writing rows." (next turn) "Wrote 50 rows to Sheet1."

## TOOL CHOICE RULES

- Spotlight search → `search_files` (NEVER `shell` with find/ls/mdfind)
- Reading text files → `read_file` with the absolute path from user's message (NEVER invent paths)
- Office automation → `applescript` (one atomic script per task)
- Opening file in specific app → `applescript` (NOT `open_file` — that uses system defaults)
- Web answers ("how do I X") → `web_lookup` (returns text), NOT `web_search` (just opens browser)
- Music/video → `play_media`, NOT vision loop
- Gmail → `check_inbox`/`search_email`/`read_email`/`send_email` (NEVER browser)
- Memory → call `save_memory` proactively when you learn user facts
- send_email is RED — only call after explicit user yes

## CONTEXT YOU RECEIVE

- `frontmost_app` — what the user is staring at
- `excel` — full Excel workbook contents when Excel is open (cells, formulas)
- `user_memory` — facts saved about the user
- Conversation history — your prior turns + tool results

USE these. If frontmost_app=Microsoft Excel and user says "fix B12", you know which workbook.

## THINKING

Use your thinking budget to reason about: target clarity, tool choice, ordering. Don't skip thinking — it's what prevents hallucinations.
"""


def build_messages(
    user_message: str,
    conversation: list[dict],
    images: list[dict] | None = None,
) -> list[dict]:
    """Build the messages array for the API call.

    `conversation` is the threaded history: a list of {role, content} dicts
    where content is either a string or a list of blocks (text, image,
    document, tool_use, tool_result).

    If `user_message` or `images` is non-empty, a new user turn is appended.
    If both are empty (e.g. when continuing after a tool_result block already
    in the conversation), nothing is appended.
    """
    messages = list(conversation)

    has_text = bool(user_message)
    has_images = bool(images)

    if not has_text and not has_images:
        return messages  # caller already added the user turn (e.g. tool_result)

    if has_images:
        content_blocks: list = []
        for img in images:
            mt = img.get("media_type", "")
            if mt == "application/pdf":
                content_blocks.append({
                    "type": "document",
                    "source": {"type": "base64", "media_type": "application/pdf", "data": img["data"]},
                })
            else:
                content_blocks.append({
                    "type": "image",
                    "source": {"type": "base64", "media_type": mt, "data": img["data"]},
                })
        if has_text:
            content_blocks.append({"type": "text", "text": user_message})
        messages.append({"role": "user", "content": content_blocks})
    else:
        messages.append({"role": "user", "content": user_message})

    return messages


def build_system_prompt(context: dict | None) -> str:
    """Append dynamic context (frontmost_app, Excel state, memory) to the
    static system prompt.
    """
    parts = [SYSTEM_PROMPT]

    if not context:
        return "\n".join(parts)

    # Mac state
    mac = context.get("mac", {}) or {}
    frontmost = context.get("frontmost_app") or mac.get("frontmost_app")
    state_lines = ["", "## CURRENT MAC STATE"]
    if frontmost:
        state_lines.append(f"frontmost_app: {frontmost}")
    if mac.get("active_document"):
        state_lines.append(f"active_document: {mac['active_document']}")
    if mac.get("browser_url"):
        state_lines.append(f"browser_url: {mac['browser_url']}")
    if mac.get("browser_title"):
        state_lines.append(f"browser_title: {mac['browser_title']}")
    if mac.get("other_open_documents"):
        state_lines.append("other_open_documents:")
        for app, doc in mac["other_open_documents"].items():
            state_lines.append(f"  {app}: {doc}")
    if mac.get("time"):
        state_lines.append(f"time: {mac['time']}")
    if mac.get("home"):
        state_lines.append(f"home_dir: {mac['home']}")
    if len(state_lines) > 2:
        parts.append("\n".join(state_lines))

    # Excel workbook contents
    excel = mac.get("excel") if isinstance(mac, dict) else None
    if excel:
        try:
            from services.prompts.context_formatter import format_context as _fmt
            excel_text = _fmt(excel)
            if excel_text:
                parts.append("\n" + excel_text)
        except Exception:
            pass

    # Memory
    memory = context.get("user_memory") or ""
    if memory:
        parts.append(f"\n## USER MEMORY\n{memory}")

    return "\n".join(parts)


def call_agent(
    user_message: str,
    conversation: list[dict],
    context: dict | None,
    images: list[dict] | None = None,
    model: str = MODEL_STANDARD,
) -> dict:
    """Run one agent turn. Returns:
      {
        "tool_uses": [{"id": str, "name": str, "input": dict}],
        "text": str,                # final text response (if any)
        "thinking": str,             # extended thinking summary
        "stop_reason": str,
        "updated_conversation": [...],  # full message list with this turn appended
        "usage": {"input_tokens": int, "output_tokens": int},
      }

    The caller (desktop agent) executes the tool_uses, builds tool_result
    blocks, and calls back with those appended to conversation.
    """
    system_prompt = build_system_prompt(context)
    messages = build_messages(user_message, conversation, images)

    # System prompt as cache-enabled block (saves ~90% on tokens for the
    # static portion within 5min window)
    system_blocks = [{
        "type": "text",
        "text": system_prompt,
        "cache_control": {"type": "ephemeral"},
    }]

    try:
        response = client.messages.create(
            model=model,
            max_tokens=4096,
            system=system_blocks,
            tools=AGENT_TOOLS,
            messages=messages,
        )
    except anthropic.APIError as e:
        logger.error("agent_v2 API error: %s", e)
        return {
            "tool_uses": [],
            "text": f"API error: {e}",
            "thinking": "",
            "stop_reason": "error",
            "updated_conversation": conversation,
            "usage": {"input_tokens": 0, "output_tokens": 0},
        }

    # Parse the response blocks
    tool_uses = []
    text_parts = []
    thinking_text = ""

    for block in response.content:
        btype = getattr(block, "type", None)
        if btype == "tool_use":
            tool_uses.append({
                "id": block.id,
                "name": block.name,
                "input": block.input,
            })
        elif btype == "text":
            text_parts.append(block.text)
        elif btype == "thinking":
            thinking_text = getattr(block, "thinking", "") or ""

    final_text = "\n".join(text_parts).strip()

    # The assistant's full content (text + tool_use blocks) is appended to the
    # conversation as-is so future turns can reference the tool_use ids.
    assistant_content = []
    for block in response.content:
        btype = getattr(block, "type", None)
        if btype == "text":
            assistant_content.append({"type": "text", "text": block.text})
        elif btype == "tool_use":
            assistant_content.append({
                "type": "tool_use",
                "id": block.id,
                "name": block.name,
                "input": block.input,
            })
        elif btype == "thinking":
            # Don't persist thinking blocks — they bloat the context
            pass

    updated_conv = messages + [{"role": "assistant", "content": assistant_content}]

    return {
        "tool_uses": tool_uses,
        "text": final_text,
        "thinking": thinking_text,
        "stop_reason": response.stop_reason,
        "updated_conversation": updated_conv,
        "usage": {
            "input_tokens": response.usage.input_tokens,
            "output_tokens": response.usage.output_tokens,
        },
    }


def append_tool_results(
    conversation: list[dict],
    results: list[dict],
) -> list[dict]:
    """Append a tool_result message to the conversation.

    `results` is a list of {tool_use_id, content, is_error?} dicts.
    """
    content_blocks = []
    for r in results:
        block = {
            "type": "tool_result",
            "tool_use_id": r["tool_use_id"],
            "content": r["content"],
        }
        if r.get("is_error"):
            block["is_error"] = True
        content_blocks.append(block)

    return conversation + [{"role": "user", "content": content_blocks}]
