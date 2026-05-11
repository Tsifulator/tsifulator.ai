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

# Default model for v2. Haiku 4.5 is fully tool-call capable, ~3× cheaper
# than Sonnet, and fast enough that multi-round loops feel snappy. Sonnet
# only kicks in when the route handler detects a complex task.
MODEL_DEFAULT = MODEL_FAST


# ── Per-model output / thinking caps ───────────────────────────────────────
# max_tokens controls the most a model can generate in a single response.
# For Haiku (no thinking) — 2048 is enough for a few tool calls + a short reply.
# For Sonnet with thinking — needs room for thinking_budget + actual output.
HAIKU_MAX_TOKENS = 2048
SONNET_MAX_TOKENS = 6000
SONNET_THINKING_BUDGET = 2000  # tokens; must be < max_tokens

# Per-model price ($/1M tokens). Cache read/write multipliers from Anthropic
# docs: cache_creation_input_tokens billed at 1.25×, cache_read_input_tokens
# at 0.1×. We track these separately when present in the usage payload.
_MODEL_PRICES = {
    MODEL_FAST: {"input": 1.0, "output": 5.0},
    MODEL_STANDARD: {"input": 3.0, "output": 15.0},
    MODEL_HEAVY: {"input": 15.0, "output": 75.0},
}


def estimate_cost(model: str, usage: dict) -> float:
    """Estimate USD cost from an Anthropic usage block.

    Handles cache_creation (1.25× input) and cache_read (0.1× input).
    """
    p = _MODEL_PRICES.get(model, _MODEL_PRICES[MODEL_FAST])
    in_rate = p["input"] / 1_000_000.0
    out_rate = p["output"] / 1_000_000.0

    inp = usage.get("input_tokens", 0)
    out = usage.get("output_tokens", 0)
    cache_write = usage.get("cache_creation_input_tokens", 0)
    cache_read = usage.get("cache_read_input_tokens", 0)

    return (
        inp * in_rate
        + out * out_rate
        + cache_write * in_rate * 1.25
        + cache_read * in_rate * 0.1
    )


# Heuristic: complex tasks that benefit from Sonnet's reasoning quality
# (multi-app orchestration, financial modeling, ambiguous data work).
# Everything else defaults to Haiku to keep credit burn low.
import re as _re_mod_v2
_COMPLEX_PATTERNS = _re_mod_v2.compile(
    r"\b(build|create|generate|construct|model|analyze|analyse|"
    r"refactor|debug|fix|explain|summari[sz]e|compare|"
    r"dupont|valuation|forecast|projection|pivot)\b",
    _re_mod_v2.IGNORECASE,
)


def pick_model(user_message: str, has_images: bool = False) -> str:
    """Pick Haiku or Sonnet based on the request shape.

    Haiku handles: search, open, play, shortcuts, memory, quick web lookups,
    one-shot AppleScript, single-step file ops.
    Sonnet handles: multi-step builds, financial modeling, anything with an
    image attachment (vision quality matters more), debugging/explaining.
    """
    if has_images:
        return MODEL_STANDARD
    if len(user_message) > 200:
        return MODEL_STANDARD
    if _COMPLEX_PATTERNS.search(user_message or ""):
        return MODEL_STANDARD
    return MODEL_FAST


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
        "name": "web_open",
        "description": "Open a web search in the user's browser (just navigation). Use when the user wants to BROWSE results themselves.",
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
        "description": "Fetch a URL and return its text content (HTML stripped). Use when you have a specific URL to read.",
        "input_schema": {
            "type": "object",
            "required": ["url"],
            "properties": {
                "url": {"type": "string"},
                "max_chars": {"type": "integer", "default": 8000},
            },
        },
    },
    # web_search is Anthropic's NATIVE server-side search tool — appended
    # below at construction time. It bypasses bot-detection on consumer
    # search engines. Use it as the primary "look something up" path.

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
    {
        "name": "create_routine",
        "description": (
            "Create a recurring background task that fires its prompt on a "
            "schedule. Use this when the user asks for something to happen "
            "automatically over time (e.g. 'every morning summarize my inbox', "
            "'every hour check market', 'every Sunday review my calendar'). "
            "The scheduler will fire `prompt` through this same agent each time."
        ),
        "input_schema": {
            "type": "object",
            "required": ["name", "prompt", "schedule"],
            "properties": {
                "name": {"type": "string", "description": "Short human label (e.g. 'Morning inbox summary')"},
                "prompt": {"type": "string", "description": "The instruction to fire each time the routine runs."},
                "schedule": {
                    "type": "string",
                    "description": (
                        "Friendly schedule. Examples: 'daily 08:00', "
                        "'weekdays 9am', 'every 30 min', 'hourly', "
                        "'mondays 10:00', 'market hours every 60 min'."
                    ),
                },
            },
        },
    },

    # ── Anthropic NATIVE server tool: web_search ────────────────────────────
    # Executed on Anthropic's servers; results come back inline as
    # `web_search_tool_result` blocks. No client execution needed.
    # Use this for ANY "look something up online" question. It bypasses the
    # bot-blocking that hits consumer search engine scrapers.
    {
        "type": "web_search_20250305",
        "name": "web_search",
        "max_uses": 5,
    },
]


# ── System prompt — kept tight to minimize per-round token cost ─────────
# Cached via ephemeral cache_control, so subsequent rounds within 5min pay
# 0.1× for these tokens. Even so: every word here gets re-billed somehow,
# so prefer rules > examples, and trust the tool schemas for details.
SYSTEM_PROMPT = """You are tsifl, a Mac automation agent controlled via tools.

# Behavior
- Iterative: think, call one or more tools, see results, think again, until done.
- Ask FIRST when ambiguous — text reply, no tool calls. Don't guess targets.
- Your final text must match what tools actually did. No fake "All set" claims.

# When to ask vs act
ASK if: target file/recipient is unclear AND no attachment, path, or context match.
ACT if: user attached a file → use it; typed a path → use it exactly; specific named target → use it.

# Tool-choice cheatsheet
- File search → search_files (NEVER shell find/ls/mdfind)
- Read text/CSV → read_file with EXACT path from user (never invent)
- Attached image/PDF → already in your context; don't read_file for it
- Office writes → applescript (one atomic script per task)
- Open file in a specific named app → applescript (open_file uses system defaults)
- Web answers → web_search (Anthropic native, has citations) — use freely
- Open a browser tab → web_open
- Read a specific URL → fetch_url
- Music/video → play_media
- Email → check_inbox/search_email/read_email/draft_email; send_email = RED, needs explicit yes
- Learn user facts → save_memory proactively

# Context provided
frontmost_app, mac.excel (workbook cells+formulas), user_memory, conversation history.
USE THEM. "Fix B12" + Excel open = you already know the cell."""


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


def prune_conversation(messages: list[dict], keep_recent: int = 6) -> list[dict]:
    """Trim verbose tool_result content from older turns to keep context lean.

    Strategy: walk from the most recent message back; the first `keep_recent`
    user/assistant exchanges are kept verbatim. For older messages, replace
    bulky tool_result content with a short stub. Plain-text messages and
    tool_use blocks are kept as-is so the conversational thread stays intact.
    """
    if len(messages) <= keep_recent * 2:
        return messages  # nothing to trim

    cutoff = len(messages) - (keep_recent * 2)
    trimmed: list[dict] = []
    for i, m in enumerate(messages):
        if i >= cutoff:
            trimmed.append(m)
            continue

        content = m.get("content")
        if isinstance(content, list):
            new_content = []
            for block in content:
                if isinstance(block, dict) and block.get("type") == "tool_result":
                    raw = block.get("content", "")
                    text = raw if isinstance(raw, str) else str(raw)
                    if len(text) > 400:
                        text = text[:300] + f" …[pruned; was {len(text)} chars]"
                    new_content.append({
                        "type": "tool_result",
                        "tool_use_id": block.get("tool_use_id", ""),
                        "content": text,
                    })
                else:
                    new_content.append(block)
            trimmed.append({"role": m["role"], "content": new_content})
        else:
            trimmed.append(m)
    return trimmed


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

    # Prune older tool_results so multi-round conversations don't bloat
    pruned_conv = prune_conversation(conversation, keep_recent=6)
    messages = build_messages(user_message, pruned_conv, images)

    # System prompt as cache-enabled block (saves ~90% on tokens for the
    # static portion within 5min window)
    system_blocks = [{
        "type": "text",
        "text": system_prompt,
        "cache_control": {"type": "ephemeral"},
    }]

    # Model-specific knobs. Sonnet/Opus get extended thinking; Haiku doesn't
    # support it (and doesn't usually need it for tool selection).
    is_sonnet_or_better = model in (MODEL_STANDARD, MODEL_HEAVY)
    max_tokens = SONNET_MAX_TOKENS if is_sonnet_or_better else HAIKU_MAX_TOKENS
    create_kwargs = {
        "model": model,
        "max_tokens": max_tokens,
        "system": system_blocks,
        "tools": AGENT_TOOLS,
        "messages": messages,
    }
    if is_sonnet_or_better:
        create_kwargs["thinking"] = {
            "type": "enabled",
            "budget_tokens": SONNET_THINKING_BUDGET,
        }

    try:
        response = client.messages.create(**create_kwargs)
    except anthropic.APIError as e:
        logger.error("agent_v2 API error: %s", e)
        # Friendly user-facing message for the common cases
        err_text = str(e).lower()
        if "credit balance" in err_text or "credit_balance" in err_text:
            msg = (
                "💳 Anthropic API credits depleted. "
                "Top up at console.anthropic.com → Billing, then try again."
            )
        elif "rate" in err_text and "limit" in err_text:
            msg = "⏱ Rate-limited by Anthropic. Wait a few seconds and retry."
        elif "overloaded" in err_text:
            msg = "🟧 Anthropic is overloaded. Try again in a moment."
        else:
            msg = f"API error: {str(e)[:200]}"
        return {
            "tool_uses": [],
            "text": msg,
            "thinking": "",
            "stop_reason": "error",
            "updated_conversation": conversation,
            "usage": {"input_tokens": 0, "output_tokens": 0},
        }

    # Parse the response blocks
    # tool_uses → CLIENT-side tools the desktop agent must execute
    # server_tool_uses → already executed by Anthropic; included for UI logging
    tool_uses = []
    server_tool_uses = []
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
        elif btype == "server_tool_use":
            server_tool_uses.append({
                "id": block.id,
                "name": block.name,
                "input": block.input,
            })
        elif btype == "text":
            text_parts.append(block.text)
        elif btype == "thinking":
            thinking_text = getattr(block, "thinking", "") or ""

    final_text = "\n".join(text_parts).strip()

    # The assistant's full content (text + tool_use + server_tool_use +
    # web_search_tool_result blocks) is appended to the conversation
    # as-is so future turns can reference the tool_use ids and see the
    # search results from native server tools.
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
        elif btype == "server_tool_use":
            # Native Anthropic server tool call (e.g. web_search). We keep
            # the block as-is so the next turn sees what was searched.
            assistant_content.append({
                "type": "server_tool_use",
                "id": block.id,
                "name": block.name,
                "input": block.input,
            })
        elif btype == "web_search_tool_result":
            # Results from the server-side search. Persist them so the
            # model can reference them in follow-up turns.
            raw_content = block.content
            # Anthropic returns the content either as a list of result
            # objects or an error object. Serialize to a list of dicts.
            if isinstance(raw_content, list):
                serialized = []
                for item in raw_content:
                    serialized.append({
                        "type": "web_search_result",
                        "url": getattr(item, "url", ""),
                        "title": getattr(item, "title", ""),
                        "encrypted_content": getattr(item, "encrypted_content", ""),
                        "page_age": getattr(item, "page_age", None),
                    })
                assistant_content.append({
                    "type": "web_search_tool_result",
                    "tool_use_id": block.tool_use_id,
                    "content": serialized,
                })
            else:
                # Error case — serialize as-is
                assistant_content.append({
                    "type": "web_search_tool_result",
                    "tool_use_id": block.tool_use_id,
                    "content": {
                        "type": "web_search_tool_result_error",
                        "error_code": getattr(raw_content, "error_code", "unknown"),
                    },
                })
        elif btype == "thinking":
            # Don't persist thinking blocks — they bloat the context
            pass

    updated_conv = messages + [{"role": "assistant", "content": assistant_content}]

    usage_block = {
        "input_tokens": response.usage.input_tokens,
        "output_tokens": response.usage.output_tokens,
        "cache_creation_input_tokens": getattr(response.usage, "cache_creation_input_tokens", 0) or 0,
        "cache_read_input_tokens": getattr(response.usage, "cache_read_input_tokens", 0) or 0,
    }
    return {
        "tool_uses": tool_uses,
        "server_tool_uses": server_tool_uses,
        "text": final_text,
        "thinking": thinking_text,
        "stop_reason": response.stop_reason,
        "updated_conversation": updated_conv,
        "usage": usage_block,
        "cost_usd": estimate_cost(model, usage_block),
        "model": model,
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
