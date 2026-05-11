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
            "Search the Mac via Spotlight (mdfind). Returns a list of absolute "
            "file paths, sorted by most-recently-modified first. "
            "Use this for ALL file searches. Never call `shell` with find, "
            "mdfind, ls, or locate. "
            "For 'open my recent <type>', call with file_type=<type> and use "
            "index 0 of the result. The sorting guarantees you get the right "
            "file without having to pick. "
            "Returns 'No files found' if nothing matches — in that case, ASK "
            "the user instead of guessing."
        ),
        "input_schema": {
            "type": "object",
            "required": ["query"],
            "properties": {
                "query": {
                    "type": "string",
                    "description": (
                        "Search term — filename or content. Use '*' (or any "
                        "single character) when you only care about file_type "
                        "and want the most recent file of that type."
                    ),
                },
                "file_type": {
                    "type": "string",
                    "enum": ["excel", "word", "ppt", "pdf", "image", "csv", "text"],
                    "description": "Filter by type. Strongly recommended for 'recent X' queries.",
                },
                "max_results": {"type": "integer", "default": 10},
            },
        },
    },
    {
        "name": "open_file",
        "description": (
            "Open a file in its macOS DEFAULT app (CSV → Numbers, PDF → Preview, "
            "DOCX → Word, etc.). "
            "DO NOT use this when the user named a specific app — use "
            "`applescript` instead. Example: user says 'open data.csv in Excel' "
            "→ applescript `tell application \"Microsoft Excel\" to open \"...\"`, "
            "NOT open_file (which would open in Numbers). "
            "Use open_file when the user said just 'open X' with no app preference."
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
            "Read a text file from disk and return its contents in the next "
            "tool_result. Use for: CSV, TSV, TXT, JSON, source code (.py, .r, "
            ".js, etc.), markdown, HTML, log files. "
            "\n\nSTRICT RULES:\n"
            "• path MUST be absolute (starts with '/'). Bare filenames are "
            "rejected; the error will tell you.\n"
            "• Use the EXACT path the user provided in their message. Do not "
            "modify, normalize, or invent paths. If the user didn't give a "
            "path, ASK for it rather than guess.\n"
            "• If the user ATTACHED an image or PDF (it appears as an image/ "
            "document block in your context), you can READ IT DIRECTLY. "
            "Calling read_file for an attached file is a hallucination — the "
            "content is already in your context window.\n"
            "\nReturns contents truncated to max_chars. If you see a "
            "'[truncated]' marker, call read_file again with a larger "
            "max_chars to see the rest."
        ),
        "input_schema": {
            "type": "object",
            "required": ["path"],
            "properties": {
                "path": {
                    "type": "string",
                    "description": "Absolute path. Must start with '/'.",
                },
                "max_chars": {
                    "type": "integer",
                    "default": 30000,
                    "description": "Max chars to return. Default is enough for typical files.",
                },
            },
        },
    },
    {
        "name": "write_file",
        "description": (
            "Write text content to a file on disk. Creates parent directories "
            "if missing. Overwrites existing files without warning. "
            "Use for: saving generated reports, CSVs, scripts, markdown notes. "
            "Path must be absolute. "
            "For OFFICE documents (.xlsx, .docx, .pptx) — use applescript "
            "instead so the content is structured correctly; write_file would "
            "produce a corrupt binary file."
        ),
        "input_schema": {
            "type": "object",
            "required": ["path", "content"],
            "properties": {
                "path": {"type": "string", "description": "Absolute path."},
                "content": {"type": "string"},
            },
        },
    },

    # ── App control ────────────────────────────────────────────────────────
    {
        "name": "open_app",
        "description": (
            "Launch or activate a Mac app by exact name. "
            "Use as a precursor to applescript when the app must be running "
            "first (rarely needed — most applescript blocks include `tell "
            "application X to activate` themselves). "
            "On its own, this just opens the app — not a complete task. "
            "If the user asked you to DO something in the app, you still need "
            "an applescript call afterward. Don't claim 'done' just because "
            "an app opened."
        ),
        "input_schema": {
            "type": "object",
            "required": ["name"],
            "properties": {
                "name": {
                    "type": "string",
                    "description": (
                        "Exact app name as it appears in /Applications, e.g. "
                        "'Microsoft Excel', 'Safari', 'Numbers', 'RStudio'."
                    ),
                },
            },
        },
    },
    {
        "name": "open_url",
        "description": (
            "Open a URL in the user's default browser. Use for: launching a "
            "specific webpage the user named, opening a SaaS dashboard, etc. "
            "For Gmail/calendar use the dedicated tools (check_inbox, etc.) "
            "instead — they're faster and don't open a tab."
        ),
        "input_schema": {
            "type": "object",
            "required": ["url"],
            "properties": {"url": {"type": "string"}},
        },
    },
    {
        "name": "applescript",
        "description": (
            "Run AppleScript. The most powerful tool — it can drive ANY Mac "
            "app: Excel, Word, PowerPoint, Numbers, Pages, Keynote, Finder, "
            "Safari, Chrome, Mail, Calendar, Reminders, and system events. "
            "\n\nUSE FOR:\n"
            "• Writing to Excel/Word/PPT cells, formulas, slides\n"
            "• Opening files in a specific app: `tell application \"Microsoft "
            "Excel\" to open \"/path/to/data.csv\"`\n"
            "• Pressing app shortcuts (Macabacus, Think-Cell) via System "
            "Events keystroke commands\n"
            "• Building entire documents in one script (preferred over "
            "multiple cell-by-cell calls)\n"
            "\nWRITE ONE ATOMIC SCRIPT per task, not many tiny ones. A whole "
            "workbook of values + formulas should be one script. "
            "\nWhen writing to Excel, target the EXACT sheet name from the "
            "mac.excel context block. If you reference a sheet name that "
            "doesn't exist, the script silently fails. "
            "\nESCAPE QUOTES carefully — AppleScript uses double quotes; if "
            "your content has double quotes, escape them as \\\". "
            "\nAfter the script runs, you'll get a string result. Read it: "
            "errors look like 'error: Microsoft Excel got an error...'."
        ),
        "input_schema": {
            "type": "object",
            "required": ["script"],
            "properties": {
                "script": {
                    "type": "string",
                    "description": "Complete AppleScript source — multiple lines OK.",
                },
            },
        },
    },
    {
        "name": "shell",
        "description": (
            "Run a read-only shell command. Returns stdout. "
            "\nDO NOT use for:\n"
            "• File search → use `search_files` (faster, Spotlight-indexed)\n"
            "• File reading → use `read_file`\n"
            "• Anything destructive (rm, mv, kill, etc.)\n"
            "• App automation → use `applescript`\n"
            "\nGOOD uses: `ls -la /tmp/some_specific_dir`, `wc -l <file>`, "
            "`date`, `whoami`, querying system info."
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
        "description": (
            "Copy text to the system clipboard. The user can then ⌘V into "
            "any app. Use when you've generated content the user will paste "
            "elsewhere (a SQL query, a code snippet, an email body to paste "
            "into Outlook, etc)."
        ),
        "input_schema": {
            "type": "object",
            "required": ["text"],
            "properties": {"text": {"type": "string"}},
        },
    },
    {
        "name": "clipboard_read",
        "description": (
            "Read the current clipboard contents. Use when the user says "
            "'use what's on my clipboard', 'paste that', or 'I just copied "
            "X — do Y with it'."
        ),
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "notify",
        "description": (
            "Show a macOS notification banner. Use sparingly — only for "
            "things the user explicitly asked to be notified about, or for "
            "background routines completing. Don't notify for normal "
            "interactive responses — your text reply is the response."
        ),
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
            "Play music or a video on a specific platform. Battle-tested per "
            "app — uses URL schemes + UI scripting + verification. "
            "FOR 'play X on spotify' OR 'play X' (no platform named) — use "
            "this tool, not a vision loop. "
            "FOR 'play X on youtube' — use this with platform='youtube'. "
            "platform='apple music' is supported but rarely needed. "
            "Returns the actual track name + artist that started playing — "
            "if it doesn't match what the user asked for, you can apologize "
            "and try again."
        ),
        "input_schema": {
            "type": "object",
            "required": ["platform", "query"],
            "properties": {
                "platform": {"type": "string", "enum": ["spotify", "youtube", "apple music"]},
                "query": {"type": "string", "description": "Song, artist, playlist, or video title."},
            },
        },
    },
    {
        "name": "web_open",
        "description": (
            "Open a web SEARCH PAGE in the user's browser. Use when the user "
            "wants to look at results themselves: 'pull up google for X', "
            "'open YouTube and find me X'. "
            "Does NOT return any text — you're just navigating. "
            "For getting actual content/answers, use `web_search` (which "
            "returns text)."
        ),
        "input_schema": {
            "type": "object",
            "required": ["query"],
            "properties": {
                "query": {"type": "string"},
                "engine": {
                    "type": "string",
                    "enum": ["google", "bing", "duckduckgo", "youtube"],
                    "default": "google",
                },
            },
        },
    },
    {
        "name": "fetch_url",
        "description": (
            "Fetch a SPECIFIC URL and return its text content (HTML stripped). "
            "Use when you have a known URL — e.g. the user shared a link, or "
            "you got a URL from a previous web_search result that you want "
            "to read in full. "
            "For exploratory 'look up X' questions, use `web_search` "
            "instead — it searches AND returns content in one call."
        ),
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
            "Export the active document from a Mac app to a file. Uses "
            "battle-tested AppleScript per app — works correctly with HFS "
            "paths, format options, etc. "
            "USE THIS for 'export this data', 'save the Numbers spreadsheet "
            "as a CSV', 'export the active document as PDF'. "
            "Do NOT roll your own AppleScript for this — the per-app script "
            "logic (Numbers uses HFS colon paths, Excel uses POSIX, etc.) is "
            "easy to get wrong."
        ),
        "input_schema": {
            "type": "object",
            "required": ["source_app", "destination"],
            "properties": {
                "source_app": {
                    "type": "string",
                    "description": "App name: 'Numbers', 'Microsoft Excel', 'Pages', 'Preview'.",
                },
                "destination": {
                    "type": "string",
                    "description": "Absolute output path. ~ in paths is expanded.",
                },
                "format": {"type": "string", "enum": ["csv", "tsv", "pdf"], "default": "csv"},
            },
        },
    },

    # ── Gmail ──────────────────────────────────────────────────────────────
    {
        "name": "check_inbox",
        "description": (
            "Fetch recent emails from the user's Gmail inbox. Returns a list "
            "with sender, subject, snippet, and message_id for each. "
            "Each result includes '[id:ABC123]' — you'll use that id with "
            "read_email or send_email's reply_to_id. "
            "User can follow up with 'read 3' or 'reply to the first one' — "
            "the message_ids are remembered for follow-up turns."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "max_results": {"type": "integer", "default": 10},
            },
        },
    },
    {
        "name": "search_email",
        "description": (
            "Search Gmail with native query syntax. Same operators as the "
            "Gmail web UI: from:, to:, subject:, is:unread, is:starred, "
            "has:attachment, newer_than:7d, older_than:1m, label:work, etc. "
            "Combine freely: 'from:dave subject:meeting newer_than:14d'. "
            "Returns the same shape as check_inbox (id, sender, subject, "
            "snippet)."
        ),
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
        "description": (
            "Read the FULL body of a specific email. Use after check_inbox or "
            "search_email returned a list — pass the [id:…] value from the "
            "result you want. "
            "Use this before drafting a reply so you can quote/reference "
            "specifics from the original."
        ),
        "input_schema": {
            "type": "object",
            "required": ["message_id"],
            "properties": {
                "message_id": {
                    "type": "string",
                    "description": "The id returned by check_inbox or search_email.",
                },
            },
        },
    },
    {
        "name": "send_email",
        "description": (
            "Send an email via Gmail. THIS IS RED RISK — it actually leaves "
            "the user's outbox and arrives in someone's inbox. "
            "\nDO NOT CALL unless the user has explicitly said 'send', 'send "
            "it', 'yes send', or similar in the same conversation. "
            "Drafting a message and asking 'should I send?' counts — wait for "
            "their confirmation. "
            "\nFor replies: ALWAYS set reply_to_id to keep the thread intact. "
            "The id comes from a prior check_inbox/search_email/read_email. "
            "\nIf you want to prepare an email but aren't sure about sending, "
            "use draft_email instead — that's yellow risk and the user can "
            "review in Gmail."
        ),
        "input_schema": {
            "type": "object",
            "required": ["to", "subject", "body"],
            "properties": {
                "to": {"type": "string", "description": "Recipient email address."},
                "subject": {"type": "string"},
                "body": {"type": "string", "description": "Plain text body."},
                "reply_to_id": {
                    "type": "string",
                    "description": "If replying, the message_id of the email being replied to. Threads the reply.",
                },
            },
        },
    },
    {
        "name": "draft_email",
        "description": (
            "Create a Gmail draft — does NOT send. The user reviews it in "
            "Gmail's drafts folder. "
            "This is the right tool for: outreach campaigns, sensitive emails, "
            "anything where the user wants a final review before sending. "
            "Also the right tool when you're not 100% confident about content."
        ),
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
    # USE THESE ONLY when no specialized tool will work: the user wants you
    # to interact with a GUI element that has no AppleScript or API path.
    # If applescript can do the job, do that instead — it's 10× more reliable.
    {
        "name": "screenshot",
        "description": (
            "Capture the screen and return it as a base64 image you can "
            "analyze in the NEXT turn. The image is returned as a tool_result; "
            "you'll see it in your conversation context. "
            "Use ONLY for: web pages with no API, custom apps with no "
            "AppleScript dictionary, anything where you need to visually "
            "identify a UI element to click. "
            "If the user is in Excel/Word/PPT/Numbers — use applescript, NOT "
            "screenshot. Office apps have full AppleScript dictionaries; "
            "vision is fragile and slow by comparison."
        ),
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
        "description": (
            "Click at screen coordinates. Coordinates come from a "
            "PREVIOUS screenshot — they map 1:1 to screen points. "
            "Always pair with a screenshot in the same workflow. "
            "Click the CENTER of the target element, not the edge. "
            "Don't fabricate coordinates — if you didn't take a screenshot, "
            "take one first."
        ),
        "input_schema": {
            "type": "object",
            "required": ["x", "y"],
            "properties": {
                "x": {"type": "integer"},
                "y": {"type": "integer"},
                "click_type": {
                    "type": "string",
                    "enum": ["left", "right", "double"],
                    "default": "left",
                },
            },
        },
    },
    {
        "name": "type_text",
        "description": (
            "Type text via simulated keystrokes into the focused field. "
            "Don't use for filling Office docs — applescript is more reliable. "
            "Use for: web forms, chat boxes, search fields after a click_at."
        ),
        "input_schema": {
            "type": "object",
            "required": ["text"],
            "properties": {"text": {"type": "string"}},
        },
    },
    {
        "name": "key_combo",
        "description": (
            "Press a keyboard shortcut. Format: '<modifier>+<modifier>+<key>'. "
            "Examples: 'cmd+c', 'cmd+shift+t', 'return', 'tab', 'escape', "
            "'cmd+option+v' (Macabacus paste format), 'alt+f4'. "
            "Modifiers: cmd, ctrl, shift, alt/option. "
            "Send to the frontmost app — make sure that's the right app "
            "before pressing."
        ),
        "input_schema": {
            "type": "object",
            "required": ["keys"],
            "properties": {"keys": {"type": "string"}},
        },
    },
    {
        "name": "scroll",
        "description": "Scroll the focused window in a direction.",
        "input_schema": {
            "type": "object",
            "required": ["direction"],
            "properties": {
                "direction": {"type": "string", "enum": ["up", "down", "left", "right"]},
                "amount": {"type": "integer", "default": 3, "description": "Wheel-tick count."},
            },
        },
    },
    {
        "name": "wait",
        "description": (
            "Wait for the UI to settle (page load, animation, etc). "
            "Use AFTER click_at/type_text when an app needs a beat to "
            "respond. Max 10 seconds — don't ask for longer."
        ),
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
            "Call this PROACTIVELY — don't ask permission, just save useful "
            "facts as a side-effect of completing the task. "
            "\nFacts worth saving: email addresses (their own, boss, "
            "professor, frequent contacts), project paths, app preferences "
            "('I use RStudio for analysis'), recurring people ('Dave is my "
            "manager'), naming conventions, default settings. "
            "\nFormat each fact as a SHORT, self-contained sentence. "
            "Good: 'Dave's email is dave@acme.com'. "
            "Bad: 'they mentioned Dave once'. "
            "\nIf you save a fact while also doing other work, save it in "
            "the same turn (parallel to your other tool calls)."
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
            "Create a persistent shortcut the user can trigger later. "
            "\nTwo flavors:\n"
            "• Slash command — `trigger='data'` → user types `/data` in tsifl\n"
            "• System hotkey — `hotkey='cmd+d'` → user presses ⌘D from ANY app\n"
            "\nUse when the user says 'set X as /name' or 'set X as cmd+D' or "
            "'make a shortcut for X'. "
            "\n`action` is the natural-language instruction tsifl will run "
            "when the shortcut fires — e.g. 'check my inbox', 'open ~/Desktop/"
            "data.csv in Excel'."
        ),
        "input_schema": {
            "type": "object",
            "required": ["trigger", "action"],
            "properties": {
                "trigger": {"type": "string", "description": "Short slug, e.g. 'data', 'inbox'."},
                "action": {"type": "string", "description": "Instruction to run when triggered."},
                "hotkey": {
                    "type": "string",
                    "description": "Optional system hotkey like 'cmd+d', 'cmd+shift+1'.",
                },
            },
        },
    },
    {
        "name": "create_routine",
        "description": (
            "Create a RECURRING background task. The scheduler will fire the "
            "`prompt` through this same agent at each scheduled time, "
            "automatically. Results appear as macOS notifications. "
            "\nUse when the user asks for something to happen automatically "
            "over time. Triggers: 'every morning', 'every hour', 'each "
            "weekday', 'every Sunday', etc. "
            "\nThe `prompt` will be executed by a future instance of you — "
            "write it as a clear self-contained instruction. "
            "\nBAD: prompt='do that thing again' (no context for future you). "
            "GOOD: prompt='check my Gmail inbox, summarize the 5 most recent "
            "unread emails, ignore newsletters and promos'."
        ),
        "input_schema": {
            "type": "object",
            "required": ["name", "prompt", "schedule"],
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Short human label, e.g. 'Morning inbox summary'.",
                },
                "prompt": {
                    "type": "string",
                    "description": "Clear self-contained instruction to fire on each run.",
                },
                "schedule": {
                    "type": "string",
                    "description": (
                        "Friendly schedule string. Examples: 'daily 08:00', "
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


# ── System prompt — instruction manual for the model ─────────────────────
# Cached via ephemeral cache_control: first request within a 5min window
# writes the cache at 1.25× cost, subsequent ones read at 0.1×. Length here
# is not the bottleneck — quality of behavioral scaffolding is.
SYSTEM_PROMPT = """You are tsifl, a Mac automation agent that acts through tools.

# How you work
You operate in a loop:
1. Read the user's message and the CURRENT MAC STATE block below.
2. Decide: is the target/file/recipient clear, or should I ASK first?
3. If asking: respond with text and NO tool calls. The loop continues when the user replies.
4. If acting: call the smallest set of tools needed for the next step.
5. Read the tool_result. Decide the next step. Repeat until the task is done.
6. When done, respond with a short factual summary of what you actually did.

# Before each tool call, think this through (silently)
- What is the user actually asking for?
- Is the target (file, app, recipient, cell) explicitly named, attached, or visible in the context block?
- If YES: which tool is the right one (see tool-choice rules below)?
- If NO: respond with a clarifying question. Do NOT guess.

# Ask vs act — concrete examples
User: "import a dataset"
  → No attached file, no path in message. ASK: "Which dataset? Drag a CSV into tsifl, or paste its full path."

User: "import /Users/me/Downloads/data.csv into a new Excel workbook"
  → Path is explicit. Call read_file(path=…) first, then applescript to write to Excel.

User: "open my latest Word doc"
  → Target is clear (latest .docx). Call search_files(file_type='word'); use the FIRST result (Spotlight sorts by recency); call applescript to open in Word.

User: "fix B12" — frontmost_app=Microsoft Excel, mac.excel.sheet=Summary
  → You know the cell. Look at the formula in the context block. If it's broken, write applescript to fix it. Don't ask.

User: "send the email"
  → No recipient, no body specified. ASK what to send and to whom.

User: drops an image showing a small table, says "import this"
  → The image is in your context as an image block. READ THE TABLE FROM THE IMAGE. Open the target app. Call applescript to write the rows. Do NOT call read_file or search_files — the data is right there.

# Tool-choice rules
- File search → `search_files`. Never `shell` with find/ls/mdfind/locate. search_files uses Spotlight and sorts by most-recently-modified.
- Read text on disk → `read_file` with the EXACT absolute path the user provided. If they didn't give a path, ASK.
- Open a file in a specific app → `applescript` (e.g. `tell application "Microsoft Excel" to open "/path/to/data.csv"`). Plain `open_file` uses macOS defaults (CSV → Numbers, not Excel).
- Office writes (Excel/Word/PPT cells, formulas, slides) → ONE `applescript` per task. Atomic.
- Web lookup → `web_search` (Anthropic native, returns real content with citations). Use freely.
- Open a search tab for the user to browse → `web_open`.
- Read a known URL → `fetch_url`.
- Music/video → `play_media`. Not vision loop.
- Gmail → `check_inbox`, `search_email`, `read_email`, `draft_email`, `send_email`. NEVER open Gmail in a browser to read mail.
- Learn user facts (their email, boss's name, project paths, preferences) → call `save_memory` proactively as part of the same turn.
- Recurring tasks → `create_routine`.

# Risk
- `send_email` is RED. Never call it without an explicit user confirmation in the SAME message ("send it", "yes send", etc.).
- All others auto-execute. Use them confidently when the target is clear.

# Attached files vs file paths — read this twice
The user can attach files in two ways:

A) ATTACHED (image, PDF) — appears as an image/document block in your input. You read its content DIRECTLY. **Never call read_file when the user has attached a file** — the content is already in your context window. Calling read_file for an attached file is a hallucination unless the user separately typed a path.

B) PATH in text — the user typed or dragged a path like `/Users/me/Downloads/data.csv`. Call `read_file(path="/Users/me/Downloads/data.csv")` with that EXACT string. Never invent a path, never modify it, never normalize it.

If neither applies and the user says "the file" or "the data" — ASK.

# After a tool returns
- Read the result before deciding the next step.
- If the result has data (file contents, search results, email body) — USE that data; don't re-fetch.
- If a tool errored — read the error. Do NOT retry with a guess. If the path was wrong, ASK the user. If the script syntax was wrong, FIX the script. If a path doesn't exist, don't search for variants — ask.
- A search_files result is a list of paths sorted by recency. For "open my recent X", use index 0. Don't pick randomly.

# Never fake completion
Your final text response must reflect what tools actually ran. Forbidden:
- "All set" / "Done" / "I've imported the data" when no write tool ran
- "Formatting applied" with only an `open_app` call
- Claiming a result that didn't appear in any tool_result

Right pattern:
- After applescript that wrote Excel cells: "Wrote 50 rows to Sheet1 of CAT_vs_Deere_DuPont_Appendix.xlsx."
- After search_files only: "Found 8 .docx files — the most recent is at /Users/.../Roosevelt_Final.docx. Want me to open it?"
- After only open_app: "Opened Excel. What should I put in it?"

# Context block (CURRENT MAC STATE)
You will receive a block below with:
- `frontmost_app` — what the user is staring at right now.
- `mac.excel` — if Excel is open: workbook name, sheet names, used range, cell values + formulas.
- `mac.other_open_documents` — other apps with open docs (Numbers, Pages, Preview).
- `user_memory` — facts saved about this user (emails, names, preferences). Use them; don't ask for things you already know.
- Recent conversation history with your prior tool calls and results.

USE THIS CONTEXT. If frontmost_app=Microsoft Excel and user says "what's in B12" — answer from the excel block; no tool call needed. If user says "email my professor" and user_memory contains "professor murphy@bc.edu" — use that address."""


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
