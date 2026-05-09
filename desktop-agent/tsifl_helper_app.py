#!/usr/bin/env python3
"""
tsifl Helper.app — macOS menu bar wrapper around the desktop agent.

Replaces the "open terminal and run python3 agent.py" flow with a
double-clickable .app that shows a menu bar icon and runs the agent in
the background. Built with rumps (PyObjC menu bar lib) + py2app for
the bundling.

Architecture:
  1. App launches → rumps app starts → menu bar icon appears
  2. Background thread starts agent.poll_and_execute()
  3. Status updates via the menu bar icon (• connected, x error)
  4. User can quit cleanly via the menu

This is the user-facing wrapper — all the actual work happens in
agent.py + excel_applescript.py, untouched.
"""

import json
import os
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

import rumps  # type: ignore[import-untyped]

# ── Shortcut-specific error log ─────────────────────────────────────────────
# stderr in a .app bundle goes to /dev/null. Real product errors need a
# persistent on-disk trail so we can debug what's failing without console
# access. Path: ~/Library/Logs/tsifl-shortcut.log
_SHORTCUT_LOG = Path.home() / "Library" / "Logs" / "tsifl-shortcut.log"


def _log_shortcut_trace(msg: str):
    """Append a non-error trace line to the shortcut log (no notification).

    Used for diagnostic events (drops received, POSTs sent, etc.) where
    we want a record but don't want to spam the user with notifications.
    """
    try:
        _SHORTCUT_LOG.parent.mkdir(parents=True, exist_ok=True)
        with _SHORTCUT_LOG.open("a", encoding="utf-8") as f:
            stamp = time.strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"{stamp} {msg}\n")
    except Exception:
        pass


def _log_shortcut_error(msg: str, exc: Exception | None = None):
    """Append an error line to the shortcut log + post a macOS notification
    so the user can SEE the failure. No more silent stderr writes."""
    try:
        _SHORTCUT_LOG.parent.mkdir(parents=True, exist_ok=True)
        with _SHORTCUT_LOG.open("a", encoding="utf-8") as f:
            stamp = time.strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"{stamp} {msg}\n")
            if exc is not None:
                import traceback
                f.write(traceback.format_exc())
                f.write("\n")
    except Exception:
        pass
    # Surface a user-visible notification with the truncated error
    try:
        import rumps as _r
        short = (str(exc) if exc else msg)[:160]
        _r.notification(title="tsifl shortcut error", subtitle=msg[:60], message=short)
    except Exception:
        pass


# Global hotkey listening: we use Carbon's RegisterEventHotKey via ctypes.
#
# Why not pynput / NSEvent? Both of those use CGEventTap under the hood,
# which requires Input Monitoring permission. That permission is keyed to
# the .app's bundle hash; every rebuild silently invalidates it (the toggle
# stays "ON" in System Settings but the listener no longer fires). Plus,
# CGEventTap is a passive observer — browsers like Chrome capture cmd+shift+t
# (reopen closed tab) before our observer ever sees it.
#
# Carbon's RegisterEventHotKey is fundamentally different: it's a system-
# level hotkey registration. Doesn't need Input Monitoring. Captures the
# keystroke BEFORE other apps see it. This is what every Mac menu bar
# utility (Raycast, Alfred, Spotlight) uses under the hood. Carbon is
# "deprecated" by Apple but still fully shipped and supported — and it'll
# outlive any of us.
import ctypes
import ctypes.util
try:
    _carbon_path = ctypes.util.find_library("Carbon")
    _carbon = ctypes.CDLL(_carbon_path) if _carbon_path else None
    _CARBON_AVAILABLE = _carbon is not None
except Exception:
    _carbon = None
    _CARBON_AVAILABLE = False

# httpx for the panel → backend /chat round-trip
try:
    import httpx as _httpx  # type: ignore[import-untyped]
    _HTTPX_AVAILABLE = True
except Exception:
    _httpx = None  # type: ignore[assignment]
    _HTTPX_AVAILABLE = False


# ── First-launch + Login Item registration ──────────────────────────────────
# Tracks whether we've already shown the first-launch dialog so we don't pester
# the user every time they reopen the app.
_FIRST_LAUNCH_MARKER = (
    Path.home() / "Library" / "Application Support" / "tsifl-helper" / ".onboarded"
)


def _is_in_login_items(app_path: str) -> bool:
    """Check if `app_path` is registered as a macOS Login Item.

    Uses AppleScript via osascript — clean, sandboxed, no external libs needed.
    Returns False on any error (we'd rather re-prompt than silently no-op).
    """
    script = (
        'tell application "System Events" to get the name of every login item'
    )
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode != 0:
            return False
        names = [n.strip() for n in result.stdout.split(",")]
        # Login Items track by app NAME, not path
        return "tsifl Helper" in names
    except Exception:
        return False


def _register_login_item(app_path: str) -> bool:
    """Add this app to macOS Login Items so it auto-starts at login.

    Note: this triggers a permission prompt on macOS Ventura+ for
    System Events automation. The user has to approve once, then it
    sticks across reboots.
    """
    script = (
        f'tell application "System Events" to make login item at end '
        f'with properties {{path:"{app_path}", hidden:false, name:"tsifl Helper"}}'
    )
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=10,
        )
        return result.returncode == 0
    except Exception:
        return False


def _maybe_show_first_launch_dialog():
    """One-time onboarding: show user a 'tsifl Helper is now running' alert
    and offer to register as a Login Item so it auto-starts.

    Idempotent — writes a marker file after the first run so subsequent
    launches skip the dialog.
    """
    if _FIRST_LAUNCH_MARKER.exists():
        return  # already onboarded

    # Discover the path of THIS .app bundle. When running from py2app, sys.argv[0]
    # points to the launcher inside Contents/MacOS — the .app is 3 levels up.
    launcher = Path(sys.argv[0]).resolve()
    if ".app/Contents/MacOS" in str(launcher):
        app_bundle = launcher.parent.parent.parent
    else:
        # Running from source (dev mode) — Login Item registration doesn't
        # apply. Still write the marker so we don't loop the dialog.
        try:
            _FIRST_LAUNCH_MARKER.parent.mkdir(parents=True, exist_ok=True)
            _FIRST_LAUNCH_MARKER.touch()
        except Exception:
            pass
        return

    response = rumps.alert(
        title="Welcome to tsifl Helper",
        message=(
            "tsifl Helper is now running in your menu bar (look for 'tsifl' "
            "near the top-right of your screen).\n\n"
            "Want it to auto-start every time you log in? You'll be asked once "
            "for permission to manage Login Items.\n\n"
            "You can change this anytime in System Settings → General → Login Items."
        ),
        ok="Auto-start at login",
        cancel="Not now",
    )

    if response == 1:  # User clicked OK / Auto-start
        if not _is_in_login_items(str(app_bundle)):
            ok = _register_login_item(str(app_bundle))
            if ok:
                rumps.notification(
                    title="tsifl Helper",
                    subtitle="Auto-start enabled",
                    message="tsifl Helper will now launch automatically when you log in.",
                )
            else:
                rumps.alert(
                    title="Couldn't enable auto-start",
                    message=(
                        "macOS blocked the request. You can add tsifl Helper "
                        "manually: System Settings → General → Login Items → +"
                    ),
                )

    # Mark as onboarded so we don't re-prompt
    try:
        _FIRST_LAUNCH_MARKER.parent.mkdir(parents=True, exist_ok=True)
        _FIRST_LAUNCH_MARKER.touch()
    except Exception:
        pass


# ── Path setup so the bundled app can find sibling modules ──────────────────
# When py2app bundles this script, sibling modules (agent.py, excel_applescript.py)
# end up in the same Resources directory. We add it to sys.path so imports work
# both during dev (running from this file) and post-bundling.
_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))


# ── Agent thread management ─────────────────────────────────────────────────

_agent_thread: threading.Thread | None = None
_agent_started_at: float | None = None
_last_error: str | None = None


def _run_agent():
    """Run the agent's main poll loop in a background thread."""
    global _last_error
    try:
        # Lazy import — runs once the menu bar app is fully initialized,
        # so any startup errors show up in the rumps log not the rumps boot
        import agent  # type: ignore[import-not-found]
        agent.main()
    except Exception as e:
        import traceback
        _last_error = f"{type(e).__name__}: {e}"
        sys.stderr.write(f"[tsifl-helper] Agent crashed: {_last_error}\n")
        sys.stderr.write(traceback.format_exc())


def _start_agent_thread():
    """Spawn the agent thread. Idempotent."""
    global _agent_thread, _agent_started_at
    if _agent_thread is not None and _agent_thread.is_alive():
        return
    _agent_thread = threading.Thread(target=_run_agent, daemon=True)
    _agent_thread.start()
    _agent_started_at = time.time()


# ── Global shortcut (Cmd+Shift+T) → floating prompt panel ───────────────────
# Pressing Cmd+Shift+T anywhere on the user's Mac pops a small modal panel
# that takes a free-text prompt, sends it to the backend's /chat endpoint
# (with the currently-frontmost app's context), and shows the reply.
#
# This is the v1 differentiator: tsifl is everywhere on your Mac, not just
# in the Excel side panel. Cmd+Shift+T from Word, from Finder, from your
# browser — all routes to the same agent.
#
# Permission requirement: pynput's global key listener requires Input
# Monitoring permission on macOS Sonoma+. First launch triggers the
# system prompt; once granted it sticks across reboots.

# Ref to the rumps app — set by TsiflHelperApp.__init__ — so the hotkey
# callback (which fires on a pynput thread) can dispatch UI to main thread.
_app_ref: "TsiflHelperApp | None" = None


def _detect_frontmost_app() -> str:
    """Return the name of the macOS frontmost application via AppleScript.

    Used to add context to the chat request so tsifl knows whether the user
    is in Excel, Word, RStudio, etc. — and routes appropriately.
    """
    try:
        result = subprocess.run(
            ["osascript", "-e",
             'tell application "System Events" to get name of first process whose frontmost is true'],
            capture_output=True, text=True, timeout=3,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
    return "unknown"


def _send_to_backend(message: str, frontmost_app: str, images: list | None = None) -> dict:
    """POST the user's prompt to the backend /chat endpoint. Synchronous —
    called from a worker thread. Returns the parsed JSON response or an
    error-shaped dict.

    Args:
        message: the user's typed text
        frontmost_app: macOS frontmost app name (e.g. "Microsoft Excel")
        images: optional list of {data, media_type, file_name} dicts —
                base64-encoded image / pdf attachments from the + button.
    """
    if not _HTTPX_AVAILABLE:
        return {"reply": "Helper is missing httpx; reinstall the bundle.",
                "actions": []}

    backend = os.getenv(
        "BACKEND_URL",
        "https://focused-solace-production-6839.up.railway.app",
    )

    # Always route through the desktop agent context now — tsifl is a
    # Mac automation agent, not an add-in chat relay.
    app = "shortcut"

    # Capture rich context about the Mac state
    try:
        from executor import get_system_context
        mac_context = get_system_context()
    except Exception:
        mac_context = {"frontmost_app": frontmost_app}

    # Load persistent user memories (preferences, facts)
    memory_context = ""
    try:
        from memory import get_memory_context
        memory_context = get_memory_context()
    except Exception:
        pass

    # Build recent conversation snippet for context
    history_snippet = ""
    if _conversation_history:
        recent = _conversation_history[-_MAX_HISTORY:]
        lines = []
        for turn in recent:
            prefix = "User" if turn["role"] == "user" else "tsifl"
            lines.append(f"{prefix}: {turn['content'][:200]}")
        history_snippet = "\n".join(lines)

    # Include last search results if available
    search_context = ""
    if _last_search_results:
        numbered = [f"  {i+1}. {p}" for i, p in enumerate(_last_search_results[:10])]
        search_context = "Recent search results:\n" + "\n".join(numbered)

    # Include recent email results so Claude can reference them by number
    email_context = ""
    if _last_email_results:
        email_lines = ["Recent email results:"]
        for i, e in enumerate(_last_email_results[:10], 1):
            sender = e.get("from", "?")
            if "<" in sender:
                sender = sender.split("<")[0].strip().strip('"')
            email_lines.append(f"  {i}. from={sender} subject={e.get('subject', '?')} id={e.get('id', '?')}")
        email_context = "\n".join(email_lines)

    body = {
        "user_id": "shortcut-anon",
        "session_id": _backend_session_id,
        "message": message,
        "context": {
            "app": app,
            "source": "global-shortcut",
            "frontmost_app": frontmost_app,
            "mac": mac_context,
            "user_memory": memory_context,
            "conversation_history": history_snippet,
            "search_results": search_context,
            "email_results": email_context,
        },
    }
    if images:
        body["images"] = images

    # Diagnostic log
    try:
        img_summary = (
            ", ".join(
                f"{(img.get('file_name') or 'pasted')}:{img.get('media_type','?')}:"
                f"{len(img.get('data','') or '')}b64chars"
                for img in (images or [])
            )
            if images else "none"
        )
        _log_shortcut_trace(
            f"POST /chat/ app={app} msg={message[:60]!r} images=[{img_summary}]"
        )
    except Exception:
        pass

    try:
        timeout = 120 if images else 60
        with _httpx.Client(timeout=timeout) as client:
            r = client.post(f"{backend.rstrip('/')}/chat/", json=body)
            if r.status_code == 200:
                data = r.json()
                # Parse plan from Claude's JSON reply (desktop agent mode)
                plan = _extract_plan_from_reply(data)
                if plan is not None:
                    data["_plan"] = plan

                # Track conversation history for multi-turn context
                reply_text = data.get("reply", "")
                _conversation_history.append({"role": "user", "content": message})
                if reply_text:
                    _conversation_history.append({"role": "assistant", "content": reply_text})
                # Trim to max
                while len(_conversation_history) > _MAX_HISTORY * 2:
                    _conversation_history.pop(0)

                return data
            return {"reply": f"Backend error ({r.status_code}): {r.text[:200]}",
                    "actions": []}
    except Exception as e:
        return {"reply": f"Could not reach tsifl: {e}", "actions": []}


def _describe_action(atype: str, payload: dict) -> str:
    """Generate a human-readable description for an action step."""
    if atype == "search_files":
        ft = payload.get("file_type", "")
        q = payload.get("query", "")
        return f"Search for {ft} files matching '{q}'" if ft else f"Search for '{q}'"
    elif atype == "open_file":
        p = payload.get("path", payload.get("file_path", ""))
        return f"Open {Path(p).name}" if p else "Open file"
    elif atype == "open_app":
        return f"Open {payload.get('app', payload.get('app_name', payload.get('name', '?')))}"
    elif atype == "open_url":
        url = payload.get("url", "")
        return f"Open {url[:60]}" if url else "Open URL"
    elif atype == "applescript":
        script = payload.get("script", payload.get("code", ""))
        return f"Run AppleScript ({script[:50]}…)" if len(script) > 50 else f"Run: {script}"
    elif atype == "shell":
        cmd = payload.get("command", "")
        return f"Run: {cmd[:60]}" if cmd else "Run shell command"
    elif atype == "clipboard_copy":
        return "Copy to clipboard"
    elif atype == "notify":
        return "Show notification"
    # Gmail actions
    elif atype == "check_inbox":
        n = payload.get("max_results", 10)
        return f"Check inbox (last {n} emails)"
    elif atype == "search_email":
        q = payload.get("query", "")
        return f"Search emails: '{q}'" if q else "Search emails"
    elif atype == "read_email":
        mid = payload.get("message_id", "")
        return f"Read email [{mid[:12]}…]" if len(mid) > 12 else f"Read email [{mid}]"
    elif atype == "send_email":
        to = payload.get("to", "?")
        return f"Send email to {to}"
    elif atype == "draft_email":
        to = payload.get("to", "?")
        return f"Draft email to {to}"
    # Screen automation
    elif atype == "screenshot":
        return "Capture screen"
    elif atype == "click_at":
        x, y = payload.get("x", "?"), payload.get("y", "?")
        return f"Click at ({x}, {y})"
    elif atype == "type_text":
        t = payload.get("text", "")
        return f"Type: {t[:40]}…" if len(t) > 40 else f"Type: {t}"
    elif atype == "key_combo":
        return f"Press {payload.get('keys', '?')}"
    elif atype == "scroll":
        return f"Scroll {payload.get('direction', 'down')}"
    elif atype == "wait":
        return f"Wait {payload.get('seconds', 1)}s"
    elif atype == "spotify_play":
        q = payload.get("query", "")
        return f"Play '{q}' on Spotify"
    return atype


def _extract_plan_from_reply(data: dict) -> list | None:
    """Extract an executable action plan from the backend response.

    Checks two sources:
    1. `actions` list (Claude's tool-use output) — convert to plan steps
    2. JSON in the reply text (fallback if Claude returned raw JSON)

    Returns a list of action dicts, or None if no actionable plan found.
    """
    # ── Source 1: Convert tool-use actions to plan steps ──────────────
    # Desktop-compatible action types that our executor can handle
    DESKTOP_TYPES = {
        "search_files", "open_file", "open_app", "open_url",
        "applescript", "shell", "clipboard_copy", "clipboard_read",
        "write_file", "notify",
        # Gmail actions
        "check_inbox", "search_email", "read_email",
        "send_email", "draft_email",
        # Screen automation (vision loop)
        "screenshot", "click_at", "type_text", "key_combo",
        "scroll", "wait",
        # App-specific fast actions
        "spotify_play",
    }

    # Risk mapping for known action types
    RISK_MAP = {
        "search_files": "green", "open_file": "green", "open_app": "green",
        "open_url": "green", "shell": "green", "clipboard_read": "green",
        "clipboard_copy": "green", "write_file": "yellow", "notify": "green",
        "applescript": "yellow",  # AppleScript can do anything
        # Gmail: reads are green, draft is yellow, send is red
        "check_inbox": "green", "search_email": "green", "read_email": "green",
        "send_email": "red", "draft_email": "yellow",
        # Screen: screenshot/scroll/wait are green, interaction is yellow
        "screenshot": "green", "scroll": "green", "wait": "green",
        "click_at": "yellow", "type_text": "yellow", "key_combo": "yellow",
        # App-specific
        "spotify_play": "green",
    }

    actions = data.get("actions") or []
    if not actions and data.get("action") and data["action"].get("type"):
        actions = [data["action"]]

    plan = []
    for act in actions:
        atype = act.get("type", "")
        payload = act.get("payload", {})
        if not isinstance(payload, dict):
            payload = {}

        if atype in DESKTOP_TYPES:
            command = json.dumps(payload) if payload else atype
            # Generate human-readable descriptions from the payload
            desc = _describe_action(atype, payload)
            plan.append({
                "type": atype,
                "description": desc,
                "command": command,
                "risk": RISK_MAP.get(atype, "yellow"),
            })

    if plan:
        return plan

    # ── Source 2: Parse JSON from reply text (fallback) ───────────────
    reply = data.get("reply", "")
    if not reply:
        return None

    import re

    # Try the entire reply as JSON
    try:
        parsed = json.loads(reply)
        if isinstance(parsed, dict) and "plan" in parsed:
            if parsed.get("reply"):
                data["reply"] = parsed["reply"]
            return parsed["plan"]
    except (json.JSONDecodeError, TypeError):
        pass

    # Try JSON inside markdown fences
    json_blocks = re.findall(r"```(?:json)?\s*\n?({.*?})\s*\n?```", reply, re.DOTALL)
    for block in json_blocks:
        try:
            parsed = json.loads(block)
            if isinstance(parsed, dict) and "plan" in parsed:
                if parsed.get("reply"):
                    data["reply"] = parsed["reply"]
                return parsed["plan"]
        except (json.JSONDecodeError, TypeError):
            pass

    return None


# ── Floating shortcut panel (Spotlight-style) ───────────────────────────────
# Custom NSPanel built via PyObjC. Replaces the previous rumps.Window modal.
# Design specs (from user):
#   - Petite, "almost like its just the box alone"
#   - Clean white background with tsifl-blue outline
#   - Small `t` logo on the left
#   - Floats above all apps, joins all spaces (Mission Control compatible)
#   - Esc to dismiss, Enter to submit
# Floats at horizontal center, ~30% from top of main screen — Spotlight-ish.

# ── Command history (up/down arrow) ─────────────────────────────────────────
_command_history: list[str] = []             # all submitted commands, newest last
_history_index: int = -1                     # -1 = not browsing; 0..N = current position
_history_saved_input: str = ""               # what user was typing before pressing up
_MAX_COMMAND_HISTORY = 50

_response_text_view: "object | None" = None # NSTextView inside the scroll view
_panel_ref: "object | None" = None         # NSPanel — module-level so it doesn't GC
_panel_input: "object | None" = None       # NSTextField inside the panel
_panel_send_btn: "object | None" = None    # send button (→)
_panel_attach_btn: "object | None" = None  # attach button (+)
_panel_response: "object | None" = None    # NSTextField for inline response below input
_panel_target: "object | None" = None      # NSObject target for input + buttons
_panel_esc_monitor: "object | None" = None # NSEvent monitor for Esc key
_panel_busy: bool = False                  # True while a backend request is in flight
_panel_thinking_timer: "object | None" = None  # rotates thinking punchlines
_panel_expanded: bool = False              # True if panel is grown to show a response.
                                            # Tracks whether subviews were shifted up so
                                            # collapse only inverts the shift when needed.
_pending_plan: list | None = None             # action plan waiting for user confirmation
_panel_session_id: int = 0                  # bumps on each show; late callbacks compare
                                            # to know if their panel-instance is still
                                            # the active one (skip mutation otherwise).

# ── Multi-turn conversation context ─────────────────────────────────────────
# Keeps a short rolling history so the user can say "open the first one"
# after a search, or "now email that to my boss" after finding a file.
# Backend session_id is persistent across panel open/close so the server
# also accumulates history on its side.
import uuid as _uuid
_backend_session_id: str = f"tsifl-desktop-{_uuid.uuid4().hex[:8]}"
_conversation_history: list[dict] = []   # [{role, content}, ...] — max ~10 turns
_MAX_HISTORY = 10
_last_search_results: list[str] = []     # file paths from most recent search
_last_email_results: list[dict] = []     # email dicts from most recent inbox/search
_last_response_text: str | None = None      # last successful reply — restored on next show
                                            # so the user doesn't lose the answer when they
                                            # ⌘⌥T-close. Cleared when a new query is sent.
_pending_images: list = []                  # list of {data: base64, media_type, file_name}
                                            # — set by attach button, sent with next submit,
                                            # cleared after each successful submit
_panel_delegate: "object | None" = None     # NSWindowDelegate that vends our custom field
                                            # editor. Held module-level so PyObjC ARC
                                            # doesn't reclaim it while the panel is live.


# ── Attachment helpers (shared by + button, drag-drop, and URL-strip fallback)
_ATTACH_EXTS = ("png", "jpg", "jpeg", "gif", "webp", "heic", "pdf")
_MEDIA_TYPES = {
    "png": "image/png",   "jpg": "image/jpeg",
    "jpeg": "image/jpeg", "gif": "image/gif",
    "webp": "image/webp", "heic": "image/heic",
    "pdf": "application/pdf",
}
# Regex pre-compile (lazy on first use to avoid import overhead at module load)
_FILE_URL_RE = None       # matches `file://...`
_PCT_ATTACH_RE = None     # matches percent-encoded filenames ending in attachment ext
_PLAIN_ATTACH_RE = None   # matches plain filenames ending in attachment ext (last-resort)
_LAST_DRAG_PATH = None    # most-recent dropped path; used by safety-net to recover
                          # the on-disk path when the input only has the rendered
                          # filename (no file:// prefix) — see _stash_dropped_path


def _resize_image_for_upload(path: str):
    """Read an image from disk, downsample if needed, return (bytes, media_type).

    Why: macOS screenshots are 5+ MB at retina resolution; the backend's
    cap is 1.4MB base64 (~1MB binary), and Claude's API tops out at 5MB.
    We resize to <1MB binary client-side so we ALWAYS fit inside both
    limits, no matter the source. Returns (None, None) if PIL is missing
    or the image can't be decoded — caller falls back to the raw bytes.

    Strategy: cap longest edge at 1568px (Anthropic's recommendation),
    then re-encode as JPEG with progressive quality reduction until the
    binary is ≤ ~700KB. JPEG over PNG because the resize is for an LLM
    vision model, not human archival — JPEG quality 80 is visually
    indistinguishable but ~5–10× smaller than PNG.
    """
    try:
        from PIL import Image as _PILImage
        import io as _io
    except Exception:
        return None, None

    try:
        img = _PILImage.open(path)
        # Flatten alpha → white background (JPEG can't carry alpha)
        if img.mode in ("RGBA", "LA"):
            bg = _PILImage.new("RGB", img.size, (255, 255, 255))
            mask = img.split()[-1]
            bg.paste(img, mask=mask)
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")

        # Cap longest edge per Anthropic vision docs.
        MAX_DIM = 1568
        if max(img.size) > MAX_DIM:
            img.thumbnail((MAX_DIM, MAX_DIM), _PILImage.LANCZOS)

        # Encode as JPEG; reduce quality until under 700KB binary
        # (≈ 930KB base64, well below the 1.4MB backend cap).
        for quality in (85, 75, 65, 55, 45):
            buf = _io.BytesIO()
            img.save(buf, format="JPEG", quality=quality, optimize=True)
            data = buf.getvalue()
            if len(data) <= 700_000:
                return data, "image/jpeg"
        return data, "image/jpeg"  # last attempt — accept whatever size
    except Exception as e:
        sys.stderr.write(f"[tsifl-helper] resize failed for {path}: {e}\n")
        return None, None


def _queue_paths_as_images(paths) -> int:
    """Read each path off disk, base64-encode, queue in `_pending_images`,
    update the +N badge on the attach button. Returns count actually added.

    Centralized so the + button, the drag-drop handler, and the URL-strip
    safety net all share one code path. Filters by extension so a stray
    text file dropped on the panel doesn't get encoded. Dedupes by
    filename so the safety-net text filter and the content-view drop
    handler don't both queue the same file. Images are resized on-disk
    if PIL is available, so a 5MB screenshot doesn't blow past the
    backend's image-size cap (1.4MB base64).
    """
    global _pending_images
    import base64 as _b64
    from pathlib import Path as _P
    added = 0
    existing_names = {item.get("file_name") for item in _pending_images}
    for raw in paths:
        try:
            path = str(raw)
            ext = _P(path).suffix.lower().lstrip(".")
            if ext not in _ATTACH_EXTS:
                continue
            name = _P(path).name
            if name in existing_names:
                continue  # dedupe — already queued

            # PDFs go through as-is (resize doesn't apply). Images get
            # downsampled to fit the backend cap. Falls back to raw
            # bytes if PIL is missing or the image can't be decoded.
            if ext == "pdf":
                with open(path, "rb") as f:
                    data = f.read()
                media_type = "application/pdf"
            else:
                resized, mt = _resize_image_for_upload(path)
                if resized is not None:
                    data = resized
                    media_type = mt
                else:
                    with open(path, "rb") as f:
                        data = f.read()
                    media_type = _MEDIA_TYPES.get(ext, "image/png")

            _pending_images.append({
                "data": _b64.b64encode(data).decode("ascii"),
                "media_type": media_type,
                "file_name": name,
            })
            existing_names.add(name)
            added += 1
        except Exception as e:
            sys.stderr.write(f"[tsifl-helper] queue image {raw} failed: {e}\n")
    if added > 0 and _panel_attach_btn is not None:
        try:
            _panel_attach_btn.setTitle_(f"+{len(_pending_images)}")
        except Exception:
            pass
    return added


def _stash_dropped_path(path: str) -> None:
    """Remember the most recent dropped file path.

    Used by the safety-net text filter: when NSTextView inserts only the
    *rendered filename* (no `file://` prefix), the regex can't recover
    the on-disk path. But we just saw the drop in our content-view
    handler, so we stash the path here and the safety net consults it.
    """
    global _LAST_DRAG_PATH
    _LAST_DRAG_PATH = path


def _strip_and_extract_paths(text: str):
    """Detect filename / URL pollution in `text`, return (cleaned_text, paths).

    Strips ONLY safe-to-strip patterns (won't false-positive on user prose):

      1. Explicit `file://...` URLs — no human types these. Decoded path
         is added to `paths` so the file gets queued as an attachment.
      2. The exact filename of `_LAST_DRAG_PATH`, in either its plain
         form ("Screenshot.png") or its URL-percent-encoded form
         ("Screenshot%202026-04-28.png"). We only strip what we KNOW
         was just dropped — avoids killing user text like "70% growth".
    """
    import re

    paths: list[str] = []
    cleaned = text

    # Pattern 1: file:// URLs — always safe to strip
    file_url_re = re.compile(r"file://\S+", re.IGNORECASE)

    def _replace_url(m):
        url = m.group(0).rstrip(".,;)")
        try:
            from urllib.parse import urlparse, unquote
            parsed = urlparse(url)
            if parsed.scheme.lower() == "file":
                p = unquote(parsed.path)
                if p:
                    paths.append(p)
        except Exception:
            pass
        return ""

    cleaned = file_url_re.sub(_replace_url, cleaned)

    # Pattern 2: match against _LAST_DRAG_PATH. Try forms in order from
    # most-specific to least-specific so we strip the longest matching
    # form first (otherwise we'd strip the filename and leave the
    # directory portion behind, e.g. "/Users/.../Desktop/").
    if _LAST_DRAG_PATH:
        from pathlib import Path as _P
        from urllib.parse import quote
        last_name = _P(_LAST_DRAG_PATH).name
        patterns: list[str] = []
        # 2a. Full on-disk path (most specific)
        patterns.append(re.escape(_LAST_DRAG_PATH))
        # 2b. URL-encoded full path
        try:
            encoded_path = quote(_LAST_DRAG_PATH, safe="/")
            if encoded_path != _LAST_DRAG_PATH:
                patterns.append(re.escape(encoded_path))
        except Exception:
            pass
        # 2c. file:// + URL-encoded path (occasionally inserted by
        # different browsers/services)
        try:
            file_url = "file://" + quote(_LAST_DRAG_PATH, safe="/")
            patterns.append(re.escape(file_url))
        except Exception:
            pass
        # 2d. Filename only (last resort)
        if last_name:
            patterns.append(re.escape(last_name))
            try:
                encoded_name = quote(last_name)
                if encoded_name != last_name:
                    patterns.append(re.escape(encoded_name))
            except Exception:
                pass

        stripped_any = False
        for pat in patterns:
            pat_re = re.compile(r"\s*" + pat + r"\s*", re.IGNORECASE)
            if pat_re.search(cleaned):
                cleaned = pat_re.sub(" ", cleaned, count=1)
                stripped_any = True
                break  # one match is enough — most-specific won
        if stripped_any and _LAST_DRAG_PATH not in paths:
            paths.append(_LAST_DRAG_PATH)

    # Tidy up: collapse multiple spaces, drop blank lines
    cleaned = "\n".join(line for line in cleaned.splitlines() if line.strip())
    cleaned = re.sub(r"  +", " ", cleaned)
    return cleaned.strip(), paths


# ── PyObjC subclasses (registered ONCE at module load) ──────────────────────
# Critical: PyObjC registers Objective-C classes with the runtime by name on
# first use. Defining a class inside a function and calling that function
# multiple times triggers a "class already registered" error on the second
# call. So we define them at module level — they're cached + reused across
# every panel rebuild.

def _define_panel_classes():
    """Define all PyObjC subclasses used by the floating panel.

    Called once on module load. Returns:
        (panel_class, target_class, content_view_class, panel_delegate_class,
         input_field_class)

    All five classes are registered with the Objective-C runtime on first
    call; subsequent panel rebuilds reuse them.
    """
    try:
        from AppKit import NSPanel, NSView, NSTextView, NSTextField
        from Foundation import NSObject
    except Exception as _ie:
        sys.stderr.write(f"[tsifl-helper] AppKit/Foundation unavailable: {_ie}\n")
        return None, None, None, None, None

    class _TsiflFloatingPanel(NSPanel):
        """Floating panel that CAN become the key window, even when
        borderless. Override required because NSPanel + borderless mask
        defaults to canBecomeKeyWindow=NO, which would block keyboard
        input from reaching the embedded text field."""

        def canBecomeKeyWindow(self):
            return True

        def canBecomeMainWindow(self):
            return False  # menu-bar utility, not the app's main window

    class _TsiflInputField(NSTextField):
        """NSTextField subclass that refuses all drag-drop at the cell
        level, so drops bubble up to the content view (which queues
        them as attachments) instead of getting inserted as URL text.

        Why subclass NSTextField (the cell-level container) and not
        NSTextView (the field editor)? Because subclassing NSTextView
        broke spacebar in v85 — key-event routing got confused. The
        cell-level NSTextField doesn't intercept keyboard input, so
        overriding only its drag methods is safe.

        unregisterDraggedTypes is also called explicitly in
        _build_floating_panel (belt-and-suspenders). This subclass
        catches anything that re-registers types dynamically.
        """

        def acceptableDragTypes(self):
            return []  # reject everything

        def draggingEntered_(self, sender):
            return 0  # NSDragOperationNone

        def draggingUpdated_(self, sender):
            return 0

        def prepareForDragOperation_(self, sender):
            return False

        def performDragOperation_(self, sender):
            return False

    class _TsiflFieldEditor(NSTextView):
        """Custom field editor used by the panel for the input NSTextField.

        Rejects ALL drag operations — belt-and-suspenders, since the
        user-reported bug was that NSTextView (the default field editor)
        accepts file URL drops and inserts them as text. We override at
        every layer (drag-types, draggingEntered, prepareFor/performDrag,
        readSelectionFromPasteboard) so the field editor cannot accept
        ANY drag-drop. File drops bubble up to `_TsiflContentView` which
        queues them as proper attachments.
        """

        def acceptableDragTypes(self):
            return []  # zero acceptable types → no drag operation

        def draggingEntered_(self, sender):
            return 0  # NSDragOperationNone

        def draggingUpdated_(self, sender):
            return 0

        def prepareForDragOperation_(self, sender):
            return False

        def performDragOperation_(self, sender):
            return False

        def readSelectionFromPasteboard_type_(self, pb, ptype):
            FILE_TYPES = (
                "NSFilenamesPboardType",
                "public.file-url",
                "public.url",
                "Apple URL pasteboard type",
            )
            if ptype in FILE_TYPES:
                return False
            try:
                return NSTextView.readSelectionFromPasteboard_type_(self, pb, ptype)
            except Exception:
                return False

        def doCommandBySelector_(self, selector):
            """Intercept up/down arrow for command history navigation."""
            sel_name = str(selector)
            if sel_name == "moveUp:" and _command_history:
                _history_navigate(-1)  # older
                return
            if sel_name == "moveDown:" and _command_history:
                _history_navigate(1)   # newer
                return
            # Default handling for everything else
            try:
                NSTextView.doCommandBySelector_(self, selector)
            except Exception:
                pass

    class _TsiflPanelDelegate(NSObject):
        """NSWindowDelegate that vends `_TsiflFieldEditor` as the field
        editor for the panel's input NSTextField.

        macOS calls `windowWillReturnFieldEditor:toObject:` whenever the
        panel needs to provide a field editor for one of its controls.
        We always return our custom subclass — the response NSTextField
        is non-editable so it never asks; only the input would.
        """

        def windowWillReturnFieldEditor_toObject_(self, window, obj):
            try:
                from Foundation import NSMakeRect
                fe = getattr(self, "_field_editor", None)
                if fe is None:
                    fe = _TsiflFieldEditor.alloc().initWithFrame_(
                        NSMakeRect(0, 0, 0, 0)
                    )
                    fe.setFieldEditor_(True)
                    # Belt-and-suspenders: also unregister all drag types
                    # at the AppKit drag-system level. Combined with the
                    # method overrides above, this gives 5+ layers of
                    # defense against file URL drops being inserted.
                    try:
                        fe.unregisterDraggedTypes()
                    except Exception:
                        pass
                    self._field_editor = fe
                return fe
            except Exception as e:
                sys.stderr.write(f"[tsifl-helper] field editor build failed: {e}\n")
                return None

    class _TsiflContentView(NSView):
        """Content view that accepts file drops anywhere on the panel.

        Files dropped here get base64-encoded and queued in
        `_pending_images` (same destination as the + button). Combined
        with the custom field editor blocking drops on the input itself,
        this gives clean attach UX: drag a file → +N badge updates,
        input stays clean.
        """

        def acceptsFirstResponder(self):
            return False  # input field owns keyboard focus

        def draggingEntered_(self, sender):
            if self._has_acceptable_drop(sender):
                return 1  # NSDragOperationCopy
            return 0      # NSDragOperationNone

        def draggingUpdated_(self, sender):
            return self.draggingEntered_(sender)

        def prepareForDragOperation_(self, sender):
            return self._has_acceptable_drop(sender)

        def performDragOperation_(self, sender):
            try:
                # Diagnostic: log what's on the pasteboard so we can see
                # why a particular drop failed to register an attachment.
                try:
                    pb = sender.draggingPasteboard()
                    types = list(pb.types() or [])
                    _log_shortcut_trace(f"drop received, pasteboard types: {types}")
                except Exception:
                    pass

                paths = self._extract_paths(sender)
                added = _queue_paths_as_images(paths)
                # Stash the most recent dropped path so the safety-net
                # text filter can recover it if the field editor somehow
                # still inserted the rendered filename.
                if paths:
                    _stash_dropped_path(paths[0])
                _log_shortcut_trace(
                    f"drop processed: paths={len(paths)} queued={added} "
                    f"pending_total={len(_pending_images)}"
                )
                return added > 0
            except Exception as e:
                _log_shortcut_error("drop failed", e)
                return False

        def _has_acceptable_drop(self, sender):
            """Cheap check (no file I/O) — does the pasteboard carry a
            type we know how to handle? Used by draggingEntered_ /
            prepareForDragOperation_ to decide whether to accept the
            drag without paying for a full _extract_paths call (which
            would stage temp files for image data).
            """
            try:
                pb = sender.draggingPasteboard()
                types = set(pb.types() or [])
                accept = {
                    "NSFilenamesPboardType",
                    "public.file-url",
                    "public.png",
                    "public.jpeg",
                    "public.tiff",
                    "public.image",
                }
                return bool(types & accept)
            except Exception:
                return False

        def _extract_paths(self, sender):
            """Pull file paths off the drag pasteboard, or stage raw image
            bytes (browser drags) to a temp file and return that path.

            Three sources, in priority order:
              1. NSFilenamesPboardType — Finder, real on-disk files.
              2. NSURL objects (file://) — same idea, modern UTI form.
              3. Raw image data on the pasteboard (PNG / JPEG / TIFF) —
                 dragging an image FROM a browser like Chrome doesn't put
                 a file path on the pasteboard; the image bytes are there
                 directly. We stage them to /tmp and return that path.
            """
            paths: list[str] = []
            try:
                pb = sender.draggingPasteboard()
            except Exception:
                return paths
            # 1. Legacy filenames type (Finder)
            try:
                files = pb.propertyListForType_("NSFilenamesPboardType")
                if files:
                    return [str(f) for f in files]
            except Exception:
                pass
            # 2. NSURL objects (modern)
            try:
                from AppKit import NSURL
                urls = pb.readObjectsForClasses_options_([NSURL], None) or []
                for u in urls:
                    p = u.path() if hasattr(u, "path") else None
                    # Only accept on-disk paths — http(s) URLs from
                    # browser drags arrive here too but can't be opened.
                    if p and not str(p).startswith(("http://", "https://")):
                        paths.append(str(p))
                if paths:
                    return paths
            except Exception:
                pass
            # 3. Raw image bytes (browser drags). Stage to a temp file.
            try:
                staged = self._stage_pasteboard_image(pb)
                if staged:
                    paths.append(staged)
            except Exception as e:
                sys.stderr.write(f"[tsifl-helper] stage pasteboard image failed: {e}\n")
            return paths

        def _stage_pasteboard_image(self, pb):
            """If the pasteboard carries raw image data (PNG/JPEG/TIFF),
            write it to a temp file and return the path. Used for the
            drag-from-browser case where there's no file URL.
            """
            import tempfile
            # Try PNG and JPEG first (no conversion needed)
            for img_type, ext in (
                ("public.png",  "png"),
                ("public.jpeg", "jpg"),
            ):
                try:
                    data = pb.dataForType_(img_type)
                    if data and data.length() > 0:
                        raw = bytes(data)
                        with tempfile.NamedTemporaryFile(
                            delete=False, suffix=f".{ext}", prefix="tsifl_drop_"
                        ) as f:
                            f.write(raw)
                            return f.name
                except Exception:
                    continue
            # TIFF needs conversion to PNG (Claude doesn't accept TIFF)
            try:
                tiff = pb.dataForType_("public.tiff")
                if tiff and tiff.length() > 0:
                    from AppKit import NSBitmapImageRep, NSBitmapImageFileTypePNG
                    rep = NSBitmapImageRep.imageRepWithData_(tiff)
                    if rep is not None:
                        png_data = rep.representationUsingType_properties_(
                            NSBitmapImageFileTypePNG, None,
                        )
                        if png_data and png_data.length() > 0:
                            raw = bytes(png_data)
                            with tempfile.NamedTemporaryFile(
                                delete=False, suffix=".png", prefix="tsifl_drop_"
                            ) as f:
                                f.write(raw)
                                return f.name
            except Exception:
                pass
            return None

    class _PanelTarget(NSObject):
        """Single target object handling Enter (input field action),
        Send button click, and Attach button click. Stateless — reads
        the current `_panel_input` ref from module globals each time."""

        def inputDidEnter_(self, sender):
            if _panel_busy:
                return  # ignore double-Enter while in flight
            text = str(sender.stringValue() or "").strip()
            # Allow empty Enter to execute a pending plan
            if text or _pending_plan:
                _panel_submit(text)

        def sendClicked_(self, sender):
            if _panel_busy or _panel_input is None:
                return
            text = str(_panel_input.stringValue() or "").strip()
            if text or _pending_plan:
                _panel_submit(text)

        def controlTextDidChange_(self, notification):
            """Primary defense against URL pollution from drag-drop.

            Fires on every user-initiated change (typing, paste, drag-
            drop). Strips file:// URLs and matches the full
            _LAST_DRAG_PATH (set when the content view sees a drop) —
            both safe to strip without false-positives on user prose.

            Fallback: if a drop bypassed the content view (drop landed
            on the input field directly, before unregisterDraggedTypes
            took effect), the path inserted as text is itself a
            recoverable file path. We detect it via the abs-path-with-
            attachment-extension regex and use it as _LAST_DRAG_PATH so
            the strip-and-extract logic kicks in.

            CRITICAL: only mutate the input when `paths` came back non-
            empty. If we mutated based on the cleaned-vs-current diff
            alone, the .strip()/whitespace-collapse in the helper would
            eat user-typed trailing spaces — making the spacebar appear
            broken whenever an attachment was queued.

            `setStringValue_` does NOT re-fire this method (only field-
            editor mutations do), so the rewrite below cannot recurse.
            """
            global _LAST_DRAG_PATH
            if _panel_input is None or _panel_busy:
                return
            try:
                current = str(_panel_input.stringValue() or "")
                if not current:
                    return

                # Fallback path-detect: if there's no _LAST_DRAG_PATH
                # but the text contains an absolute path ending in an
                # attachment extension AND that path exists on disk,
                # promote it to _LAST_DRAG_PATH so the strip logic
                # below has something to anchor on.
                if not _LAST_DRAG_PATH and "/" in current:
                    import re as _re
                    ext_alt = "|".join(_ATTACH_EXTS)
                    m = _re.search(
                        r"(/[^\n]+?\.(?:" + ext_alt + r"))",
                        current, _re.IGNORECASE,
                    )
                    if m:
                        from pathlib import Path as _P
                        candidate = m.group(1).rstrip(" .,;)")
                        if _P(candidate).exists():
                            _LAST_DRAG_PATH = candidate
                            _queue_paths_as_images([candidate])

                # Fast-path bail: no file:// URL AND no recent drop.
                if "file://" not in current.lower() and not _LAST_DRAG_PATH:
                    return

                cleaned, paths = _strip_and_extract_paths(current)
                # Only apply the cleaned text if we actually matched a
                # file URL or filename pattern (paths non-empty). Pure
                # typing produces no match → leave the user's text alone
                # (including any trailing whitespace they're mid-typing).
                if paths and cleaned != current:
                    _panel_input.setStringValue_(cleaned)
                if paths:
                    _queue_paths_as_images(paths)
            except Exception:
                pass

        def attachClicked_(self, sender):
            """Attach button. Behavior depends on current state:

              - "+"  (no attachments) → click opens the file picker
              - "+N" (attachments queued) → click CLEARS all attachments

            Binary toggle. Picked over the v86 ⌥-click approach because
            ⌥-click was undiscoverable + NSApp.currentEvent() doesn't
            always carry the modifier flags at action-fire time.

            To attach multiple files: NSOpenPanel allows multi-select
            (Cmd+click in the picker), or drag-drop additional files
            onto the panel — both still work.
            """
            global _pending_images, _LAST_DRAG_PATH

            # If anything's queued, the click means "clear" (delete UX)
            if _pending_images:
                _pending_images = []
                _LAST_DRAG_PATH = None
                if _panel_attach_btn is not None:
                    try:
                        _panel_attach_btn.setTitle_("+")
                    except Exception:
                        pass
                if _panel_input is not None:
                    try:
                        _panel_input.becomeFirstResponder()
                    except Exception:
                        pass
                return

            # Otherwise: open file picker → encode + queue
            try:
                from AppKit import NSOpenPanel
                panel = NSOpenPanel.openPanel()
                panel.setAllowsMultipleSelection_(True)
                panel.setCanChooseFiles_(True)
                panel.setCanChooseDirectories_(False)
                panel.setAllowedFileTypes_(list(_ATTACH_EXTS))
                panel.setMessage_("Attach an image or PDF for tsifl")
                panel.setPrompt_("Attach")
                if panel.runModal() != 1:  # 1 = NSModalResponseOK
                    return
                urls = panel.URLs() or []
                paths = [u.path() for u in urls if u.path()]
                _queue_paths_as_images(paths)
                if _panel_input is not None:
                    try:
                        _panel_input.becomeFirstResponder()
                    except Exception:
                        pass
            except Exception as e:
                _log_shortcut_error("attach failed", e)

    return (_TsiflFloatingPanel, _PanelTarget, _TsiflContentView,
            _TsiflPanelDelegate, _TsiflInputField)


(
    _TsiflFloatingPanel,
    _PanelTargetClass,
    _TsiflContentViewClass,
    _TsiflPanelDelegateClass,
    _TsiflInputFieldClass,
) = _define_panel_classes()


# Thinking punchlines shown in the response area while waiting for backend.
# Reused / inspired by the Excel addin's typing animation; analyst-flavored
# but never overstays its welcome (rotates every ~1.5s).
_THINKING_LINES = (
    "Thinking…",
    "Reading the room…",
    "Crunching numbers like it's bonus szn…",
    "Asking VLOOKUP for advice…",
    "Counting rows like a back-office analyst at 2am…",
    "Sharpening pencils…",
    "Pulling up the comp set…",
    "Triple-checking the math…",
    "Reading the room (round 2)…",
    "The intern's still researching…",
)


def _build_floating_panel():
    """Construct the NSPanel + content view + input field + inline buttons.

    Layout (collapsed = input row only, ~52px tall):
      ┌──────────────────────────────────────────────────────┐
      │ t  Ask tsifl anything…                       +   →   │
      └──────────────────────────────────────────────────────┘

    When a request is in flight or a response is shown, a second row
    appears below the input — the panel grows to fit the response.

    Returns (panel, input_field, send_button, attach_button, response_field).
    """
    try:
        from AppKit import (
            NSPanel, NSView, NSTextField, NSImageView, NSImage, NSColor,
            NSFont, NSButton, NSBackingStoreBuffered, NSScreen,
            NSFloatingWindowLevel, NSLineBreakByTruncatingTail,
        )
    except Exception as _ie:
        sys.stderr.write(f"[tsifl-helper] AppKit import failed: {_ie}\n")
        raise
    from Foundation import NSMakeRect

    # Raw style/collection-behavior values (avoids depending on PyObjC
    # constant exports which are flaky on some Python 3.14 builds).
    NSWindowStyleMaskBorderless = 0
    NSWindowStyleMaskNonactivatingPanel = 1 << 7   # 128 — panel can be key
                                                    # without activating the
                                                    # underlying app
    NSWindowCollectionBehaviorCanJoinAllSpaces = 1 << 0
    NSWindowCollectionBehaviorFullScreenAuxiliary = 1 << 8
    NSFocusRingTypeNone = 1

    # NSPanel subclass is defined at module level (see _define_panel_classes
    # above) — defining it here would re-register the class on every call
    # and crash on the second show.

    # Compact pill — wider than tall, just one row in collapsed state.
    width = 540.0
    INPUT_ROW_HEIGHT = 52.0
    height = INPUT_ROW_HEIGHT  # starts collapsed

    screen_frame = NSScreen.mainScreen().frame()
    x = (screen_frame.size.width - width) / 2.0
    y = screen_frame.size.height * 0.72 - height / 2.0

    # NonactivatingPanel mask + borderless gives us a key-able floating panel
    # that doesn't steal app activation from whatever the user was working in.
    style_mask = NSWindowStyleMaskBorderless | NSWindowStyleMaskNonactivatingPanel

    panel = _TsiflFloatingPanel.alloc().initWithContentRect_styleMask_backing_defer_(
        NSMakeRect(x, y, width, height),
        style_mask,
        NSBackingStoreBuffered,
        False,
    )

    panel.setLevel_(NSFloatingWindowLevel)
    panel.setCollectionBehavior_(
        NSWindowCollectionBehaviorCanJoinAllSpaces
        | NSWindowCollectionBehaviorFullScreenAuxiliary
    )
    panel.setHidesOnDeactivate_(False)
    panel.setHasShadow_(True)
    panel.setOpaque_(False)
    panel.setBackgroundColor_(NSColor.clearColor())
    try:
        panel.setMovableByWindowBackground_(True)
    except Exception:
        pass

    blue = NSColor.colorWithCalibratedRed_green_blue_alpha_(
        13.0 / 255, 94.0 / 255, 175.0 / 255, 1.0
    )
    ink = NSColor.colorWithCalibratedRed_green_blue_alpha_(
        10.0 / 255, 14.0 / 255, 26.0 / 255, 1.0
    )
    muted = NSColor.colorWithCalibratedRed_green_blue_alpha_(
        100.0 / 255, 116.0 / 255, 139.0 / 255, 1.0
    )

    # Content view = white pill with blue border + rounded corners.
    # Use the drag-drop-aware subclass (`_TsiflContentView`) so dragging an
    # image file onto the panel queues it as an attachment instead of
    # letting the default field editor insert the URL as text.
    view_cls = _TsiflContentViewClass or NSView
    content = view_cls.alloc().initWithFrame_(NSMakeRect(0, 0, width, height))
    content.setWantsLayer_(True)
    layer = content.layer()
    layer.setBackgroundColor_(NSColor.whiteColor().CGColor())
    layer.setCornerRadius_(14.0)
    layer.setBorderWidth_(1.5)
    layer.setBorderColor_(blue.CGColor())
    # Register for file URL + raw image drag types. Three categories:
    #   - File paths from Finder (NSFilenamesPboardType, public.file-url)
    #   - Raw image bytes from browser drags (public.png, public.jpeg,
    #     public.tiff) — Chrome/Safari put the image data directly on the
    #     pasteboard with no on-disk file
    #   - Generic image UTI (public.image) — covers HEIC, GIF, etc.
    # The drop handler stages raw bytes to /tmp via _stage_pasteboard_image
    # so the downstream queue logic sees a real path either way.
    if _TsiflContentViewClass is not None:
        try:
            content.registerForDraggedTypes_([
                "NSFilenamesPboardType",
                "public.file-url",
                "public.png",
                "public.jpeg",
                "public.tiff",
                "public.image",
            ])
        except Exception as e:
            sys.stderr.write(f"[tsifl-helper] register drag types failed: {e}\n")
    panel.setContentView_(content)

    # ── Top row: logo + input + attach + send ────────────────────────────
    # All laid out absolutely. Y-coords use the input-row coordinate space.

    # 1. Blue `t` logo on the left. Uses icon_blue.png if available,
    #    else falls back to the black template icon (tinted blue via NSImageView).
    logo_size = 20.0
    logo_x = 14.0
    logo_y = (INPUT_ROW_HEIGHT - logo_size) / 2.0

    blue_icon_path = None
    try:
        # Look for icon_blue.png next to the bundled icon.png
        ic = TsiflHelperApp._resolve_icon_path()
        if ic:
            from pathlib import Path as _P
            cand = _P(ic).parent / "icon_blue.png"
            if cand.exists():
                blue_icon_path = str(cand)
    except Exception:
        pass

    icon_path_to_use = blue_icon_path or TsiflHelperApp._resolve_icon_path()
    if icon_path_to_use:
        try:
            logo_image = NSImage.alloc().initWithContentsOfFile_(icon_path_to_use)
            if logo_image is not None:
                logo_view = NSImageView.alloc().initWithFrame_(
                    NSMakeRect(logo_x, logo_y, logo_size, logo_size)
                )
                logo_view.setImage_(logo_image)
                content.addSubview_(logo_view)
        except Exception as e:
            sys.stderr.write(f"[tsifl-helper] logo load failed: {e}\n")

    # 2. Send button (rightmost)
    btn_size = 26.0
    send_x = width - btn_size - 10.0
    send_y = (INPUT_ROW_HEIGHT - btn_size) / 2.0
    send_btn = NSButton.alloc().initWithFrame_(
        NSMakeRect(send_x, send_y, btn_size, btn_size)
    )
    send_btn.setBordered_(False)
    send_btn.setTitle_("→")
    send_btn.setFont_(NSFont.boldSystemFontOfSize_(16.0))
    try:
        send_btn.setContentTintColor_(blue)
    except Exception:
        pass
    content.addSubview_(send_btn)

    # 3. Attach button (left of send)
    attach_x = send_x - btn_size - 2.0
    attach_y = (INPUT_ROW_HEIGHT - btn_size) / 2.0
    attach_btn = NSButton.alloc().initWithFrame_(
        NSMakeRect(attach_x, attach_y, btn_size, btn_size)
    )
    attach_btn.setBordered_(False)
    # Restore +N badge if attachments were queued in a prior panel session
    # (user closed the panel without submitting, then reopened).
    attach_btn.setTitle_(f"+{len(_pending_images)}" if _pending_images else "+")
    attach_btn.setFont_(NSFont.systemFontOfSize_(18.0))
    try:
        attach_btn.setContentTintColor_(muted)
    except Exception:
        pass
    # Tooltip explains the binary-toggle behavior. Hover lingers for
    # ~1s before showing in macOS by default.
    try:
        attach_btn.setToolTip_(
            "Click + to attach an image or PDF.  Click +N again to clear."
        )
    except Exception:
        pass
    content.addSubview_(attach_btn)

    # 4. Input field (between logo and attach button). Use the drag-
    # rejecting NSTextField subclass so file drops bubble up to the
    # content view (which queues them as attachments) instead of being
    # inserted as URL text inside the input.
    input_x = logo_x + logo_size + 12.0
    input_width = attach_x - input_x - 8.0
    input_height = 24.0
    input_y = (INPUT_ROW_HEIGHT - input_height) / 2.0
    input_cls = _TsiflInputFieldClass or NSTextField
    input_field = input_cls.alloc().initWithFrame_(
        NSMakeRect(input_x, input_y, input_width, input_height)
    )
    # Belt-and-suspenders: also unregister any drag types AppKit might
    # default-register on the field. Combined with the subclass overrides
    # above, this gives us 6 layers of "no drops on the input".
    try:
        input_field.unregisterDraggedTypes()
    except Exception:
        pass
    input_field.setBezeled_(False)
    input_field.setBordered_(False)
    input_field.setDrawsBackground_(False)
    input_field.setFocusRingType_(NSFocusRingTypeNone)
    input_field.setFont_(NSFont.systemFontOfSize_(15.0))
    input_field.setTextColor_(ink)
    input_field.setStringValue_("")
    input_field.setPlaceholderString_("Ask tsifl anything…")
    content.addSubview_(input_field)

    # ── Response area (multi-line wrapping, hidden until first submit) ──
    # Multi-line text so the full reply is visible, not truncated to one
    # line. Width-bounded to the panel's content area; height grows
    # dynamically when we measure the reply (see _panel_show_response).
    # ── Scrollable response area (NSTextView inside NSScrollView) ──────
    from AppKit import NSScrollView, NSTextView, NSBorderType
    NSNoBorder = 0

    scroll_view = NSScrollView.alloc().initWithFrame_(NSMakeRect(0, 0, 0, 0))
    scroll_view.setHasVerticalScroller_(True)
    scroll_view.setHasHorizontalScroller_(False)
    scroll_view.setBorderType_(NSNoBorder)
    scroll_view.setDrawsBackground_(False)
    scroll_view.setAutohidesScrollers_(True)
    scroll_view.setHidden_(True)

    text_view = NSTextView.alloc().initWithFrame_(NSMakeRect(0, 0, _PANEL_WIDTH - 36, 0))
    text_view.setEditable_(False)
    text_view.setSelectable_(True)
    text_view.setFont_(NSFont.systemFontOfSize_(13.0))
    text_view.setTextColor_(muted)
    text_view.setDrawsBackground_(False)
    text_view.setTextContainerInset_((0, 0))
    # Word wrap to container width
    text_view.textContainer().setWidthTracksTextView_(True)
    text_view.setHorizontallyResizable_(False)
    text_view.setVerticallyResizable_(True)

    scroll_view.setDocumentView_(text_view)
    content.addSubview_(scroll_view)

    # Store the inner text view in a module-level ref — can't set attrs
    # on ObjC objects, so we use a global instead.
    global _response_text_view
    _response_text_view = text_view
    response_field = scroll_view

    return panel, input_field, send_btn, attach_btn, response_field


def _make_panel_target():
    """Allocate an instance of the module-level _PanelTargetClass.

    The class itself is registered once at module load (see
    _define_panel_classes). This function just creates a new instance,
    which is safe to call any number of times.
    """
    if _PanelTargetClass is None:
        return None
    return _PanelTargetClass.alloc().init()


# ── Panel size constants ────────────────────────────────────────────────
_INPUT_ROW_HEIGHT = 52.0
_PANEL_WIDTH = 600.0
_RESPONSE_PADDING_X = 18.0   # horizontal padding inside the response area
_RESPONSE_PADDING_TOP = 8.0  # gap between response and input row
_RESPONSE_PADDING_BOTTOM = 12.0
_MIN_RESPONSE_HEIGHT = 36.0   # always show at least one full line + padding
_MAX_RESPONSE_HEIGHT = 500.0  # taller cap so long replies are readable


def _measure_text_height(text: str, width: float, font_size: float = 13.0) -> float:
    """Measure how tall the wrapped text needs to render at the given width.

    Uses AppKit's NSAttributedString boundingRectWithSize so it matches
    exactly what NSTextField will draw — no off-by-one approximations.
    """
    try:
        from AppKit import NSAttributedString, NSFontAttributeName, NSFont
        from Foundation import NSMakeSize
        # NSStringDrawingUsesLineFragmentOrigin = 1, NSStringDrawingUsesFontLeading = 16
        opts = (1 << 0) | (1 << 4)
        font = NSFont.systemFontOfSize_(font_size)
        attrs = {NSFontAttributeName: font}
        attr_str = NSAttributedString.alloc().initWithString_attributes_(text, attrs)
        rect = attr_str.boundingRectWithSize_options_(
            NSMakeSize(width, 9999.0), opts,
        )
        import math
        return math.ceil(rect.size.height) + 4.0  # tiny pad
    except Exception:
        # Fallback: rough estimate, ~14px per ~85 chars per line
        import math
        chars_per_line = max(1, int(width / 7.0))
        lines = max(1, math.ceil(len(text) / chars_per_line))
        return lines * 18.0


def _panel_show_response(text: str):
    """Resize the panel to show the full response below the input row.

    Multi-line, word-wrapped, height computed from actual text size so a
    short reply gets a small panel and a long reply gets a tall one
    (capped at _MAX_RESPONSE_HEIGHT so we don't fill the screen on a
    pathological 10000-char reply).

    Top edge stays fixed; panel grows downward in window coords.
    """
    global _panel_expanded
    if _panel_ref is None or _panel_response is None:
        return
    try:
        from Foundation import NSMakeRect

        # Compute height needed for the response text at the panel's width.
        text_width = _PANEL_WIDTH - (2 * _RESPONSE_PADDING_X)
        measured = _measure_text_height(text, text_width)
        response_height = max(_MIN_RESPONSE_HEIGHT, min(measured, _MAX_RESPONSE_HEIGHT))

        new_panel_height = (
            _INPUT_ROW_HEIGHT
            + _RESPONSE_PADDING_TOP
            + response_height
            + _RESPONSE_PADDING_BOTTOM
        )

        cur_frame = _panel_ref.frame()
        # Anchor the TOP edge: keep top_y fixed, grow downward (lower origin Y)
        top_y = cur_frame.origin.y + cur_frame.size.height
        new_y = top_y - new_panel_height

        # Compute how much taller we just got vs. before — used to know how
        # far to shift the input-row subviews up so they stay at the top.
        delta = new_panel_height - cur_frame.size.height

        _panel_ref.setFrame_display_(
            NSMakeRect(cur_frame.origin.x, new_y, _PANEL_WIDTH, new_panel_height),
            True,
        )

        # Shift input-row subviews up by `delta` so they sit at the top
        # of the now-taller panel (in window coords, y increases upward).
        content = _panel_ref.contentView()
        for sub in content.subviews():
            if sub == _panel_response:
                continue
            f = sub.frame()
            sub.setFrameOrigin_((f.origin.x, f.origin.y + delta))

        # Position the response scroll view in the bottom area
        _panel_response.setFrame_(NSMakeRect(
            _RESPONSE_PADDING_X,
            _RESPONSE_PADDING_BOTTOM,
            text_width,
            response_height,
        ))
        # Set text on the inner NSTextView (stored as _tsifl_text_view)
        tv = _response_text_view
        if tv is not None:
            tv.setString_(text)
            # Scroll to top on new content
            tv.scrollRangeToVisible_((0, 0))
        else:
            # Fallback for old-style NSTextField
            _panel_response.setStringValue_(text)
        _panel_response.setHidden_(False)
        _panel_expanded = True
    except Exception as e:
        sys.stderr.write(f"[tsifl-helper] show response failed: {e}\n")


def _panel_collapse_response():
    """Hide response + shrink panel back to input-row only.

    No-op if not currently expanded — prevents subview shift corruption
    when a late callback hits after dismiss.
    """
    global _panel_expanded
    if _panel_ref is None or _panel_response is None:
        return
    def _clear_response_text():
        tv = _response_text_view
        if tv is not None:
            tv.setString_("")
        else:
            _panel_response.setStringValue_("")

    if not _panel_expanded:
        try:
            _panel_response.setHidden_(True)
            _clear_response_text()
        except Exception:
            pass
        return
    try:
        from Foundation import NSMakeRect
        _panel_response.setHidden_(True)
        _clear_response_text()

        cur_frame = _panel_ref.frame()
        top_y = cur_frame.origin.y + cur_frame.size.height
        new_y = top_y - _INPUT_ROW_HEIGHT
        delta = _INPUT_ROW_HEIGHT - cur_frame.size.height  # negative

        content = _panel_ref.contentView()
        for sub in content.subviews():
            if sub == _panel_response:
                continue
            f = sub.frame()
            sub.setFrameOrigin_((f.origin.x, f.origin.y + delta))

        _panel_ref.setFrame_display_(
            NSMakeRect(cur_frame.origin.x, new_y, _PANEL_WIDTH, _INPUT_ROW_HEIGHT),
            True,
        )
        _panel_expanded = False
    except Exception as e:
        sys.stderr.write(f"[tsifl-helper] collapse failed: {e}\n")


def _panel_dismiss():
    """Hide and DESTROY the panel. We rebuild fresh on next show — that's
    the simplest way to guarantee no carry-over state corruption.

    Why destroy + rebuild instead of reuse? Caching the NSPanel between
    shows turned out fragile: orderOut + frame mutations + subview shifts +
    late callbacks from in-flight requests all combined in subtle ways
    that left the panel blank/invisible/un-typeable on subsequent opens.
    Rebuild from scratch is ~5ms and bug-proof.
    """
    global _panel_busy, _panel_thinking_timer, _pending_plan
    global _panel_ref, _panel_input, _panel_send_btn, _panel_attach_btn
    global _panel_response, _panel_expanded, _vision_loop_active
    _pending_plan = None  # cancel any pending action plan
    # Kill any running vision loop — Esc is the emergency stop
    if _vision_loop_active:
        _vision_loop_active = False
        sys.stderr.write("[tsifl-helper] VISION LOOP KILLED by Esc/dismiss\n")
        try:
            rumps.notification(title="tsifl", subtitle="", message="⏹️ Vision loop stopped.")
        except Exception:
            pass

    if _panel_thinking_timer is not None:
        try: _panel_thinking_timer.stop()
        except Exception: pass
        _panel_thinking_timer = None
    _panel_busy = False

    if _panel_ref is not None:
        try:
            _panel_ref.orderOut_(None)
        except Exception:
            pass

    # Drop all panel-instance refs so the underlying NSPanel + subviews
    # are deallocated by ObjC ARC. _panel_target and _panel_esc_monitor
    # stay (they're registered once per app lifetime, not per show).
    _panel_ref = None
    _panel_input = None
    _panel_send_btn = None
    _panel_attach_btn = None
    _panel_response = None
    global _response_text_view
    _response_text_view = None
    _panel_expanded = False


def _history_navigate(direction: int):
    """Navigate command history. direction: -1 = older, +1 = newer."""
    global _history_index, _history_saved_input
    if not _command_history or _panel_input is None:
        return

    if _history_index == -1:
        # Starting to browse — save current input
        try:
            _history_saved_input = str(_panel_input.stringValue() or "")
        except Exception:
            _history_saved_input = ""
        _history_index = len(_command_history)

    _history_index += direction

    if _history_index < 0:
        _history_index = 0
    elif _history_index >= len(_command_history):
        # Past the newest → restore saved input
        _history_index = -1
        try:
            _panel_input.setStringValue_(_history_saved_input)
        except Exception:
            pass
        return

    try:
        _panel_input.setStringValue_(_command_history[_history_index])
    except Exception:
        pass


def _panel_is_visible() -> bool:
    """True if the panel is currently on-screen and key/visible."""
    if _panel_ref is None:
        return False
    try:
        return bool(_panel_ref.isVisible())
    except Exception:
        return False


_MAX_VISION_ROUNDS = 12  # safety cap — don't loop forever
_vision_loop_active = False  # True while a vision loop is running
                              # Set to False by Esc or error → kills the loop


def _vision_loop_continue(original_task: str, screenshot_b64: str,
                          last_reply: str, session_id: int,
                          prev_results: list, round_num: int = 1):
    """Send a screenshot back to Claude and execute the next round of actions.

    This is the core of screen automation: screenshot → Claude → actions → repeat.
    Runs on a background thread (panel stays hidden the whole time).

    Safety:
    - Checks `_vision_loop_active` before each round — Esc kills it
    - Stops on any action failure (don't blindly continue after errors)
    - Hard cap at _MAX_VISION_ROUNDS
    - Panel hidden for the duration — uses macOS notifications for status
    """
    global _vision_loop_active

    # ── Safety checks ────────────────────────────────────────────────
    if not _vision_loop_active:
        sys.stderr.write("[tsifl-helper] vision loop killed (flag cleared)\n")
        _vision_loop_done("⏹️ Vision loop stopped.", session_id)
        return

    if round_num > _MAX_VISION_ROUNDS:
        _vision_loop_active = False
        _vision_loop_done(f"⏹️ Stopped after {_MAX_VISION_ROUNDS} rounds (safety limit).", session_id)
        return

    # Check for errors in previous round — stop if anything failed
    for a in prev_results:
        if a.type != "screenshot" and not a.success:
            _vision_loop_active = False
            _vision_loop_done(
                f"⏹️ Stopped: {a.description} failed\n   ⚠️ {a.error or 'unknown error'}",
                session_id,
            )
            return

    # ── Hide panel + remember which app was active ─────────────────
    # We must re-activate the target app before clicking, because
    # hiding the panel shifts macOS focus to the Python process.
    frontmost_before = _detect_frontmost_app()
    # Skip our own process names
    if frontmost_before.lower() in ("python", "tsifl", "tsifl helper"):
        frontmost_before = ""

    def _hide_panel_for_vision():
        if _panel_ref is not None:
            try:
                _panel_ref.orderOut_(None)
            except Exception:
                pass
    try:
        from PyObjCTools.AppHelper import callAfter
        callAfter(_hide_panel_for_vision)
    except Exception:
        pass
    time.sleep(0.15)  # quick hide — panel is already hidden after round 1

    # ── Build follow-up message ──────────────────────────────────────
    action_summary = []
    for a in prev_results:
        if a.type != "screenshot":
            icon = "✅" if a.success else "❌"
            action_summary.append(f"{icon} {a.description}")
    actions_text = "\n".join(action_summary) if action_summary else "(screenshot only)"

    follow_up = (
        f"[VISION ROUND {round_num}] I executed the actions. Results:\n"
        f"{actions_text}\n\n"
        f"Here is a screenshot of the current screen (coordinates map 1:1 to screen points for click_at). "
        f"Original task: \"{original_task}\"\n"
        f"Look carefully at the screenshot. Identify the exact UI element you need to interact with. "
        f"For buttons/icons, click the CENTER of the element — not the text label next to it. "
        f"Continue with the next steps. When done, return an empty plan."
    )

    images = [{
        "data": screenshot_b64,
        "media_type": "image/jpeg",
        "file_name": f"screen_round_{round_num}.jpg",
    }]

    sys.stderr.write(f"[tsifl-helper] vision round {round_num}: sending screenshot to Claude\n")

    # ── Kill check before network call ───────────────────────────────
    if not _vision_loop_active:
        _vision_loop_done("⏹️ Vision loop stopped.", session_id)
        return

    frontmost = _detect_frontmost_app()
    result = _send_to_backend(follow_up, frontmost, images=images)
    plan = result.get("_plan")
    reply_text = (result.get("reply") or "").strip()

    # Log what Claude returned for diagnostics
    _log_shortcut_trace(
        f"vision round {round_num} response: reply={reply_text[:100]!r} "
        f"plan={plan is not None} plan_len={len(plan) if plan else 0} "
        f"actions_raw={len(result.get('actions', []))} "
        f"model={result.get('model_used', '?')}"
    )

    # ── Kill check after network call ────────────────────────────────
    if not _vision_loop_active:
        _vision_loop_done("⏹️ Vision loop stopped.", session_id)
        return

    if not plan or not isinstance(plan, list) or len(plan) == 0:
        _vision_loop_active = False
        _log_shortcut_trace(f"vision loop ended: no plan returned. reply={reply_text[:200]!r}")
        _vision_loop_done(reply_text or last_reply or "✅ Done.", session_id)
        return

    # ── Execute the next round of actions ────────────────────────────
    from executor import Action, Risk, execute_plan
    actions = []
    for step in plan:
        actions.append(Action(
            type=step.get("type", "shell"),
            description=step.get("description", ""),
            command=step.get("command", ""),
            risk=Risk(step.get("risk", "green")),
        ))

    # Log what we're about to do
    step_names = [s.get("type", "?") for s in plan]
    sys.stderr.write(f"[tsifl-helper] vision round {round_num} executing: {step_names}\n")

    # ── Re-activate the target app before clicking ───────────────
    # Hiding the panel shifts macOS focus to Python. We need the
    # target app frontmost so CGEvent clicks land on its windows.
    has_clicks = any(a.type in ("click_at", "type_text", "key_combo") for a in actions)
    if has_clicks:
        target_app = frontmost_before
        # Also check if any action opens an app — use that instead
        for a in actions:
            if a.type == "open_app":
                try:
                    cmd_data = json.loads(a.command) if a.command.startswith("{") else {}
                    target_app = cmd_data.get("app", cmd_data.get("app_name", cmd_data.get("name", "")))
                except Exception:
                    pass
        if target_app and target_app.lower() not in ("python", "tsifl", "tsifl helper"):
            try:
                from executor import run_applescript
                run_applescript(f'tell application "{target_app}" to activate', timeout=3)
                time.sleep(0.2)
                sys.stderr.write(f"[tsifl-helper] activated {target_app!r} before clicking\n")
            except Exception:
                pass

    results = execute_plan(actions)

    # ── Kill check after execution ───────────────────────────────────
    if not _vision_loop_active:
        _vision_loop_done("⏹️ Vision loop stopped.", session_id)
        return

    # Check for screenshot in results → continue loop
    screenshot_b64_next = None
    for a in results:
        if a.type == "screenshot" and a.success and a.result:
            screenshot_b64_next = a.result

    if screenshot_b64_next:
        _vision_loop_continue(
            original_task, screenshot_b64_next, reply_text,
            session_id, results, round_num + 1,
        )
    else:
        # No screenshot requested — Claude is done
        _vision_loop_active = False
        lines = []
        if reply_text:
            lines.append(reply_text)
            lines.append("")
        for a in results:
            if a.type != "screenshot":
                icon = "✅" if a.success else "❌"
                lines.append(f"{icon} {a.description}")
                if a.result and len(a.result) < 300:
                    lines.append(f"   → {a.result}")
        _vision_loop_done("\n".join(lines) or "✅ Done.", session_id)


def _vision_loop_done(summary: str, session_id: int):
    """Finish the vision loop — notify user and make panel available.

    The panel does NOT force itself back on screen (true background mode).
    Instead a macOS notification tells the user it's done — they can
    summon the panel with ⌘⌥T if they want details.
    """
    global _vision_loop_active
    _vision_loop_active = False

    def _finish():
        global _panel_busy, _last_response_text
        _panel_busy = False
        _last_response_text = summary
        # Update panel content (but don't force it visible)
        if _panel_is_visible() and _panel_session_id == session_id:
            _panel_show_response(summary)
            if _panel_input is not None:
                try:
                    _panel_input.setEditable_(True)
                    _panel_input.setPlaceholderString_("Ask tsifl anything…")
                except Exception:
                    pass
        # Always notify — user is doing other things
        try:
            # Truncate for notification
            short = summary.split("\n")[0][:120] if summary else "Task complete"
            rumps.notification(title="tsifl ✅", subtitle="Done", message=short)
        except Exception:
            pass
    try:
        from PyObjCTools.AppHelper import callAfter
        callAfter(_finish)
    except Exception:
        pass


def _panel_submit(text: str):
    """User pressed Enter or clicked Send. Don't dismiss — keep panel open,
    show 'thinking…' inline, fire backend, then show the reply inline.

    If there's a pending action plan and the user presses Enter with an
    empty input, execute the plan. If they type something new, treat it
    as a new query (cancelling the pending plan).

    If the user dismisses the panel mid-request, the request still runs
    to completion in the background; the result surfaces as a macOS
    notification rather than being silently dropped.
    """
    global _panel_busy, _panel_thinking_timer, _pending_plan

    global _pending_images
    if _panel_busy:
        return

    # ── Plan execution: empty Enter with a pending plan = EXECUTE ─────
    if not text and _pending_plan:
        plan_to_run = _pending_plan
        _pending_plan = None
        _panel_busy = True
        _panel_show_response("Executing…")

        def _run_plan():
            try:
                from executor import Action, Risk, execute_plan
                actions = []
                for step in plan_to_run:
                    actions.append(Action(
                        type=step.get("type", "shell"),
                        description=step.get("description", ""),
                        command=step.get("command", ""),
                        risk=Risk(step.get("risk", "yellow")),
                    ))
                results = execute_plan(actions)

                # Build result summary
                lines = []
                for a in results:
                    icon = "✅" if a.success else "❌"
                    lines.append(f"{icon} {a.description}")
                    if a.result and len(a.result) < 200:
                        lines.append(f"   → {a.result}")
                    if a.error:
                        lines.append(f"   ⚠️ {a.error}")
                summary = "\n".join(lines)
            except Exception as e:
                summary = f"Execution failed: {e}"

            def _show_results():
                global _panel_busy, _last_response_text
                _panel_busy = False
                _last_response_text = summary
                if _panel_is_visible():
                    _panel_show_response(summary)
                    if _panel_input is not None:
                        try:
                            _panel_input.setEditable_(True)
                            _panel_input.setPlaceholderString_("Ask tsifl anything…")
                            _panel_input.becomeFirstResponder()
                        except Exception:
                            pass
                else:
                    try:
                        rumps.notification(title="tsifl", subtitle="Done", message=summary[:200])
                    except Exception:
                        pass

            try:
                from PyObjCTools.AppHelper import callAfter
                callAfter(_show_results)
            except Exception:
                pass

        threading.Thread(target=_run_plan, daemon=True).start()
        return

    # New query typed — cancel any pending plan
    _pending_plan = None

    # ── Local memory commands (no backend needed) ────────────────────
    try:
        from memory import check_memory_intent
        mem_response = check_memory_intent(text)
        if mem_response is not None:
            _panel_show_response(mem_response)
            if _panel_input is not None:
                try:
                    _panel_input.setStringValue_("")
                    _panel_input.setPlaceholderString_("Ask tsifl anything…")
                except Exception:
                    pass
            return
    except Exception:
        pass

    # ── Quick "open N" for search results (no backend needed) ────────
    if _last_search_results:
        import re as _re
        _open_match = _re.match(
            r"^(?:open|launch|show)\s+(?:the\s+)?(?:(\d+)(?:st|nd|rd|th)?|first|second|third|last)$",
            text.strip(), _re.IGNORECASE
        )
        if _open_match:
            ordinals = {"first": 1, "second": 2, "third": 3, "last": -1}
            num_str = _open_match.group(1)
            if num_str:
                idx = int(num_str)
            else:
                word = text.strip().split()[-1].lower()
                idx = ordinals.get(word, 1)

            if idx == -1:
                idx = len(_last_search_results)
            if 1 <= idx <= len(_last_search_results):
                path = _last_search_results[idx - 1]
                try:
                    from executor import open_file
                    ok, msg = open_file(path)
                    result_text = f"✅ Opened {os.path.basename(path)}" if ok else f"❌ {msg}"
                except Exception as e:
                    result_text = f"❌ {e}"
                _panel_show_response(result_text)
                if _panel_input is not None:
                    try:
                        _panel_input.setStringValue_("")
                        _panel_input.setPlaceholderString_("Ask tsifl anything…")
                    except Exception:
                        pass
                return

    # Save to command history (skip duplicates of last command)
    global _history_index
    if not _command_history or _command_history[-1] != text:
        _command_history.append(text)
        if len(_command_history) > _MAX_COMMAND_HISTORY:
            _command_history.pop(0)
    _history_index = -1  # reset browse position

    _panel_busy = True

    my_session = _panel_session_id

    # Snapshot any pending image attachments before clearing the queue.
    images_to_send = list(_pending_images)
    _pending_images = []
    if _panel_attach_btn is not None and images_to_send:
        try:
            _panel_attach_btn.setTitle_("+")  # reset badge
        except Exception:
            pass

    # Background mode: hide panel immediately, work runs in background.
    # Results come back as macOS notifications.
    # Must use callAfter for reliable main-thread UI updates.
    try:
        from PyObjCTools.AppHelper import callAfter as _ca_hide
        def _do_hide():
            if _panel_ref is not None:
                _panel_ref.orderOut_(None)
            try:
                rumps.notification(title="tsifl", subtitle="On it…", message=text[:100])
            except Exception:
                pass
        _ca_hide(_do_hide)
    except Exception:
        # Fallback: try direct hide
        if _panel_ref is not None:
            try:
                _panel_ref.orderOut_(None)
            except Exception:
                pass

    def _do_request():
        """Runs on a daemon thread — must NOT touch UI directly. UI updates
        get dispatched back to the main thread via PyObjCTools.AppHelper.callAfter
        which is the only reliable cross-thread main-runloop dispatch in PyObjC."""
        frontmost = _detect_frontmost_app()
        try:
            result = _send_to_backend(text, frontmost, images=images_to_send or None)
        except Exception as e:
            result = {"reply": f"Request failed: {e}", "actions": []}

        def _on_main():
          try:
            global _panel_busy, _panel_thinking_timer
            # Always reset busy + stop rotation regardless of session state
            _panel_busy = False
            if _panel_thinking_timer is not None:
                try: _panel_thinking_timer.stop()
                except Exception: pass
                _panel_thinking_timer = None

            reply_text = (result.get("reply") or "(no reply)").strip()
            plan = result.get("_plan")  # parsed action plan from desktop agent mode

            sys.stderr.write(f"[tsifl-helper] _on_main: reply={reply_text[:80]!r} plan={plan is not None} visible={_panel_is_visible()} session_match={_panel_session_id == my_session}\n")

            # ── Plan mode ────────────────────────────────────────────────
            if plan and isinstance(plan, list) and len(plan) > 0:
                # Auto-execute green + yellow actions. Only RED requires
                # explicit confirmation (irreversible: send email, delete,
                # purchase). This allows the vision loop to flow without
                # stopping on every click/type action.
                has_red = any(
                    step.get("risk", "yellow") == "red" for step in plan
                )

                if not has_red:
                    # ── AUTO-EXECUTE: green/yellow actions run immediately ─
                    sys.stderr.write(f"[tsifl-helper] no-red plan ({len(plan)} steps) → auto-executing\n")
                    # Show brief feedback then hide panel — work runs in background
                    if _panel_session_id == my_session and _panel_is_visible():
                        _panel_show_response(reply_text if reply_text and reply_text != "(no reply)" else "Running…")
                    _panel_busy = True
                    # Auto-dismiss panel after 1.5s — actions run in background
                    def _auto_dismiss_panel():
                        time.sleep(1.5)
                        try:
                            from PyObjCTools.AppHelper import callAfter as _ca2
                            _ca2(lambda: _panel_ref.orderOut_(None) if _panel_ref else None)
                        except Exception:
                            pass
                    threading.Thread(target=_auto_dismiss_panel, daemon=True).start()

                    def _auto_run():
                        try:
                            from executor import Action, Risk, execute_plan
                            actions = []
                            for step in plan:
                                actions.append(Action(
                                    type=step.get("type", "shell"),
                                    description=step.get("description", ""),
                                    command=step.get("command", ""),
                                    risk=Risk(step.get("risk", "green")),
                                ))
                            sys.stderr.write(f"[tsifl-helper] executing {len(actions)} actions: {[a.type for a in actions]}\n")
                            results = execute_plan(actions)
                            for a in results:
                                sys.stderr.write(f"[tsifl-helper]   {a.type}: {'✅' if a.success else '❌'} {a.result or a.error or ''}\n")

                            # Track search results for multi-turn ("open the first one")
                            global _last_search_results, _last_email_results
                            for a in results:
                                if a.type == "search_files" and a.success and a.result:
                                    paths = [p.strip() for p in a.result.split("\n") if p.strip()]
                                    _last_search_results = paths

                            # Track email results for multi-turn ("read 3", "reply to the first one")
                            for a in results:
                                if a.type in ("check_inbox", "search_email") and a.success and a.result:
                                    import re as _re_mod
                                    ids = _re_mod.findall(r"\[id:([^\]]+)\]", a.result)
                                    if ids:
                                        _last_email_results = [{"id": mid} for mid in ids]

                            # ── Vision loop: if a screenshot was captured, feed it
                            # back to Claude so it can see what happened and decide
                            # the next step. This is the core of screen automation.
                            screenshot_b64 = None
                            for a in results:
                                if a.type == "screenshot" and a.success and a.result:
                                    screenshot_b64 = a.result

                            # Check if Claude asked for continuation
                            wants_continue = result.get("continue", False)
                            if isinstance(result.get("_raw_reply"), dict):
                                wants_continue = result["_raw_reply"].get("continue", wants_continue)

                            if screenshot_b64 and (wants_continue or any(a.type == "screenshot" for a in results)):
                                # Vision loop: hide panel immediately → runs in background
                                global _vision_loop_active
                                _vision_loop_active = True
                                sys.stderr.write(f"[tsifl-helper] vision loop STARTING (background) for: {text[:60]!r}\n")
                                try:
                                    from PyObjCTools.AppHelper import callAfter as _ca
                                    def _hide_now():
                                        if _panel_ref is not None:
                                            _panel_ref.orderOut_(None)
                                    _ca(_hide_now)
                                except Exception:
                                    pass
                                try:
                                    rumps.notification(
                                        title="tsifl 👁️",
                                        subtitle="Working in background…",
                                        message=f"Task: {text[:80]}",
                                    )
                                except Exception:
                                    pass
                                _vision_loop_continue(
                                    text, screenshot_b64, reply_text,
                                    my_session, results
                                )
                                return  # _vision_loop_continue handles the rest

                            # Build result summary
                            lines = []
                            if reply_text and reply_text != "(no reply)":
                                lines.append(reply_text)
                                lines.append("")
                            for a in results:
                                icon = "✅" if a.success else "❌"
                                lines.append(f"{icon} {a.description}")
                                if a.type == "search_files" and a.success and a.result:
                                    paths = [p.strip() for p in a.result.split("\n") if p.strip()]
                                    for idx, p in enumerate(paths[:10], 1):
                                        name = os.path.basename(p)
                                        folder = os.path.dirname(p).replace(str(Path.home()), "~")
                                        lines.append(f"   {idx}. {name}")
                                        lines.append(f"      {folder}")
                                elif a.type in ("check_inbox", "search_email", "read_email") and a.success and a.result:
                                    lines.append(a.result)
                                elif a.type != "screenshot" and a.result and len(a.result) < 300:
                                    lines.append(f"   → {a.result}")
                                if a.error:
                                    lines.append(f"   ⚠️ {a.error}")
                            if _last_search_results:
                                lines.append("")
                                lines.append("Say \"open 1\" or \"open the first one\" to open a result")
                            summary = "\n".join(lines)
                        except Exception as e:
                            sys.stderr.write(f"[tsifl-helper] _auto_run CRASHED: {e}\n")
                            import traceback; traceback.print_exc(file=sys.stderr)
                            summary = f"Execution failed: {e}"

                        def _show():
                            global _panel_busy, _last_response_text
                            _panel_busy = False
                            _last_response_text = summary
                            if _panel_is_visible():
                                _panel_show_response(summary)
                                if _panel_input is not None:
                                    try:
                                        _panel_input.setEditable_(True)
                                        _panel_input.setPlaceholderString_("Ask tsifl anything…")
                                        _panel_input.becomeFirstResponder()
                                    except Exception:
                                        pass
                            else:
                                try:
                                    rumps.notification(title="tsifl", subtitle="Done", message=summary[:200])
                                except Exception:
                                    pass

                        try:
                            from PyObjCTools.AppHelper import callAfter
                            callAfter(_show)
                        except Exception:
                            pass

                    threading.Thread(target=_auto_run, daemon=True).start()
                    return

                # ── CONFIRM: yellow/red actions need user approval ────────
                risk_icons = {"green": "🟢", "yellow": "🟡", "red": "🔴"}
                lines = [reply_text, ""]
                for i, step in enumerate(plan, 1):
                    icon = risk_icons.get(step.get("risk", "yellow"), "🟡")
                    desc = step.get("description", step.get("type", "?"))
                    lines.append(f"  {icon} {i}. {desc}")
                lines.append("")
                lines.append("Press Enter to execute  •  Esc to cancel")
                display = "\n".join(lines)

                global _last_response_text
                _last_response_text = display

                # RED plan needs confirmation — re-show panel if hidden
                sys.stderr.write(f"[tsifl-helper] showing plan confirmation ({len(plan)} steps)\n")
                if _panel_ref is not None and not _panel_is_visible():
                    try:
                        _panel_ref.makeKeyAndOrderFront_(None)
                    except Exception:
                        pass
                    try:
                        rumps.notification(title="tsifl ⚠️", subtitle="Needs confirmation", message=text[:80])
                    except Exception:
                        pass
                _panel_show_response(display)
                # Store plan for execution on confirm
                global _pending_plan
                _pending_plan = plan
                if _panel_input is not None:
                    try:
                        _panel_input.setEditable_(True)
                        _panel_input.setStringValue_("")
                        _panel_input.setPlaceholderString_("Enter = execute  •  type to ask something else")
                        _panel_input.becomeFirstResponder()
                    except Exception:
                        pass
                return

            # ── Normal reply mode (no plan) ──────────────────────────────
            display = reply_text
            if len(display) > 1500:
                display = display[:1497] + "…"

            _last_response_text = display

            if _panel_session_id == my_session and _panel_is_visible():
                _panel_show_response(display)
                if _panel_input is not None:
                    try:
                        _panel_input.setEditable_(True)
                        _panel_input.setPlaceholderString_("Ask tsifl anything…")
                        _panel_input.becomeFirstResponder()
                    except Exception:
                        pass
            else:
                try:
                    rumps.notification(
                        title="tsifl",
                        subtitle=text[:60],
                        message=display,
                    )
                except Exception:
                    pass
          except Exception as e:
            sys.stderr.write(f"[tsifl-helper] _on_main CRASHED: {type(e).__name__}: {e}\n")
            import traceback
            traceback.print_exc(file=sys.stderr)

        # PyObjCTools.AppHelper.callAfter is the canonical
        # "dispatch this to the main run loop from any thread" primitive.
        # Without this, rumps.Timer on a background thread NEVER fires
        # because NSTimer only fires on the run loop of its creating thread.
        try:
            from PyObjCTools.AppHelper import callAfter
            callAfter(_on_main)
        except Exception as e:
            sys.stderr.write(f"[tsifl-helper] main-thread dispatch failed: {e}\n")

    threading.Thread(target=_do_request, daemon=True).start()


def _show_shortcut_panel():
    """Toggle the floating panel. If visible → dismiss + destroy. Otherwise
    build a brand new panel and show it.

    Rebuild-fresh-each-show is the foolproof pattern. Caching the panel
    led to subtle layout / focus / stale-callback bugs that were hard to
    fully eliminate. Rebuilding takes ~5ms — fast enough that the user
    can't tell, and guarantees zero state carryover.
    """
    global _panel_ref, _panel_input, _panel_send_btn, _panel_attach_btn
    global _panel_response, _panel_target, _panel_esc_monitor
    global _panel_session_id, _panel_expanded

    # TOGGLE: hotkey while panel is up → close (and destroy)
    sys.stderr.write(f"[tsifl-helper] _show_shortcut_panel: visible={_panel_is_visible()} ref={_panel_ref is not None}\n")
    if _panel_is_visible():
        sys.stderr.write("[tsifl-helper] panel visible → dismissing (toggle)\n")
        _panel_dismiss()
        return

    # Always destroy any lingering state before rebuilding (defensive)
    _panel_dismiss()

    # Kill any vision loop that's still running in the background.
    # Without this, a loop started from a prior session keeps clicking
    # even after the user opens a new panel.
    global _vision_loop_active
    if _vision_loop_active:
        _vision_loop_active = False
        sys.stderr.write("[tsifl-helper] VISION LOOP KILLED by new panel open\n")

    # Bump session ID — any in-flight request from a prior show will see
    # the change and route its result to a notification instead of the
    # (now-different) panel.
    _panel_session_id += 1
    _panel_expanded = False

    try:
        # Build fresh panel + subviews
        (
            _panel_ref,
            _panel_input,
            _panel_send_btn,
            _panel_attach_btn,
            _panel_response,
        ) = _build_floating_panel()

        # Build the target the FIRST time only (it's a stateless NSObject
        # that just calls back into module-level _panel_input via globals;
        # we don't need a fresh one each show).
        if _panel_target is None:
            _panel_target = _make_panel_target()

        # Wire input field's Enter + button clicks every time (new subviews
        # need new target/action wiring).
        _panel_input.setTarget_(_panel_target)
        _panel_input.setAction_("inputDidEnter:")
        _panel_send_btn.setTarget_(_panel_target)
        _panel_send_btn.setAction_("sendClicked:")
        _panel_attach_btn.setTarget_(_panel_target)
        _panel_attach_btn.setAction_("attachClicked:")

        # Also wire the input as the target's delegate, so
        # `controlTextDidChange_` fires on every user-initiated edit.
        # That's the primary URL-pollution defense: detects file:// or
        # percent-encoded filename text in the input on every keystroke,
        # strips it, and auto-queues the corresponding file as a proper
        # attachment via _LAST_DRAG_PATH.
        try:
            _panel_input.setDelegate_(_panel_target)
        except Exception as e:
            sys.stderr.write(f"[tsifl-helper] input delegate attach failed: {e}\n")

        # NOTE: v85 tried to install a custom NSWindowDelegate (vending a
        # _TsiflFieldEditor that rejects all file URL drops) — this broke
        # the spacebar and other typing input on the field. Reason most
        # likely: NSTextView subclass needs more init wiring than just
        # setFieldEditor_:True for keyboard event routing. Reverted to
        # default field editor + relying on the controlTextDidChange_
        # safety net + content-view drag handler. The class definitions
        # are kept (in `_define_panel_classes`) for a future re-enable
        # once we can verify they don't break typing.

        # Esc monitor: register once for the app's lifetime. The handler
        # dismisses whatever the current panel is.
        if _panel_esc_monitor is None:
            try:
                from AppKit import NSEvent
                NS_EVENT_MASK_KEY_DOWN = 1 << 10
                ESC_KEYCODE = 53

                def _esc_handler(event):
                    if event.keyCode() == ESC_KEYCODE and _panel_ref is not None:
                        _panel_dismiss()
                        return None  # consume the Esc
                    return event

                _panel_esc_monitor = NSEvent.addLocalMonitorForEventsMatchingMask_handler_(
                    NS_EVENT_MASK_KEY_DOWN, _esc_handler,
                )
            except Exception as e:
                sys.stderr.write(f"[tsifl-helper] Esc monitor attach failed: {e}\n")

        # Activate process (so we can be the key window) + show + focus
        from AppKit import NSApp
        try:
            NSApp().activateIgnoringOtherApps_(True)
        except Exception:
            pass

        # Restore previous response if we have one — so the user doesn't
        # lose the last answer when ⌘⌥T-toggling. Does NOT count as
        # _panel_busy; user can immediately type a new query and submit
        # while seeing the prior answer.
        if _last_response_text and not _panel_busy:
            try:
                _panel_show_response(_last_response_text)
            except Exception as e:
                sys.stderr.write(f"[tsifl-helper] restore prev response failed: {e}\n")

        _panel_ref.orderFrontRegardless()
        _panel_ref.makeKeyWindow()
        _panel_ref.makeFirstResponder_(_panel_input)
        _panel_input.selectText_(None)

        # Source-level URL-pollution fix: strip drag-type registrations
        # on the actual field editor in use. We hit it three ways for
        # max coverage:
        #   1. NSControl.currentEditor() — the field editor right now,
        #      after makeFirstResponder. Most direct.
        #   2. NSWindow.fieldEditor:forObject: — the panel's shared
        #      cached field editor. May or may not be the same instance.
        #   3. The NSTextField itself was already unregistered in
        #      _build_floating_panel.
        # Why no NSTextView subclass here? Because that broke spacebar
        # in v85 (key-event routing got confused). Just calling
        # unregisterDraggedTypes on the default editor is safe.
        seen_editors = set()
        for desc, get in (
            ("currentEditor",       lambda: _panel_input.currentEditor()),
            ("fieldEditor_forObject", lambda: _panel_ref.fieldEditor_forObject_(True, _panel_input)),
        ):
            try:
                fe = get()
                if fe is not None and id(fe) not in seen_editors:
                    seen_editors.add(id(fe))
                    fe.unregisterDraggedTypes()
                    _log_shortcut_trace(f"field editor unregistered ({desc}): {fe}")
            except Exception as e:
                sys.stderr.write(
                    f"[tsifl-helper] field editor unregister via {desc} failed: {e}\n"
                )
    except Exception as e:
        # No more silent fallback to rumps.Window — surface the actual
        # error so we can debug. Writes to ~/Library/Logs/tsifl-shortcut.log
        # AND posts a macOS notification with the error message.
        _log_shortcut_error("panel show failed", e)


def _on_shortcut_pressed():
    """pynput callback for Cmd+Shift+T. Runs on a pynput thread; dispatches
    the panel UI to the rumps main thread via NSObject.performSelector."""
    if _app_ref is None:
        return
    try:
        # rumps timers are the cleanest way to bounce work onto the main
        # thread from an arbitrary background thread. We schedule a one-shot
        # timer that fires immediately, runs the panel, then cancels itself.
        timer_holder: list = [None]
        def _run_once(timer):
            timer.stop()
            try:
                _show_shortcut_panel()
            except Exception as e:
                sys.stderr.write(f"[tsifl-helper] shortcut panel error: {e}\n")
        t = rumps.Timer(_run_once, 0.05)
        timer_holder[0] = t
        t.start()
    except Exception as e:
        sys.stderr.write(f"[tsifl-helper] shortcut dispatch failed: {e}\n")


# ── Carbon RegisterEventHotKey via ctypes ───────────────────────────────────
# Constants from Carbon's headers. These are stable Apple-published values.
# References:
#   - kEventClassKeyboard = 'keyb' = 0x6B657962
#   - kEventHotKeyPressed = 5
#   - kEventParamDirectObject = '----' = 0x2D2D2D2D
#   - typeEventHotKeyID = 'hkid' = 0x686B6964
#   - cmdKey   = 0x100   (256 — bit 8)
#   - shiftKey = 0x200   (512 — bit 9)
#   - optionKey = 0x800
#   - kVK_ANSI_T = 0x11  (the 't' key's virtual keycode)
_KEVENT_CLASS_KEYBOARD = 0x6B657962
_KEVENT_HOTKEY_PRESSED = 5
_KEVENT_PARAM_DIRECT = 0x2D2D2D2D
_TYPE_EVENT_HOTKEY_ID = 0x686B6964
_CMD_KEY = 0x100
_SHIFT_KEY = 0x200
_OPTION_KEY = 0x800
_VK_T = 0x11  # 'T' key virtual keycode

# Carbon C structs as ctypes
class _EventHotKeyID(ctypes.Structure):
    _fields_ = [
        ("signature", ctypes.c_uint32),
        ("id",        ctypes.c_uint32),
    ]

class _EventTypeSpec(ctypes.Structure):
    _fields_ = [
        ("eventClass", ctypes.c_uint32),
        ("eventKind",  ctypes.c_uint32),
    ]

# Module-level holders so GC doesn't drop the registered handler/hotkey
_carbon_handler_ref: "object | None" = None
_carbon_hotkey_ref: "object | None" = None
_carbon_callback_holder: "object | None" = None


def _carbon_hotkey_callback(next_handler, event, user_data):
    """C callback invoked by the Carbon event manager when our hotkey fires.

    Signature: OSStatus (EventHandlerCallRef, EventRef, void*)
    We receive the event ref but don't need to inspect it — we registered
    only ONE hotkey, so any invocation here means ⌘⌥T was pressed.

    Returning 0 (noErr) tells Carbon we handled the event. The keystroke
    is consumed and doesn't propagate to whatever app is frontmost.
    """
    try:
        _on_shortcut_pressed()
    except Exception as e:
        sys.stderr.write(f"[tsifl-helper] hotkey callback error: {e}\n")
    return 0  # noErr


def _start_hotkey_listener() -> bool:
    """Register Cmd+Option+T as a system-level global hotkey via Carbon.

    Why ⌘⌥T (cmd+option+t) instead of ⌘⇧T?
    Browsers (Chrome, Safari, Arc, Firefox) all use ⌘⇧T as 'reopen closed
    tab' and capture it at the application level before any global hotkey
    layer sees it. Switching to ⌘⌥T avoids that conflict — it's not bound
    in any major Mac app I'm aware of.

    Carbon's RegisterEventHotKey works at the OS event-dispatch level, NOT
    via Input Monitoring. So this should fire regardless of whether the
    .app is signed, rebuilt, or relocated.
    """
    global _carbon_handler_ref, _carbon_hotkey_ref, _carbon_callback_holder

    if not _CARBON_AVAILABLE:
        sys.stderr.write("[tsifl-helper] Carbon framework not available\n")
        return False

    try:
        # 1. Build an EventHandlerUPP from our Python callback.
        # CFUNCTYPE signature: returns int32, takes (void*, void*, void*)
        _CallbackType = ctypes.CFUNCTYPE(
            ctypes.c_int32, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
        )
        callback_c = _CallbackType(_carbon_hotkey_callback)
        # MUST be held module-level — if GC'd, Carbon segfaults when the
        # hotkey fires.
        _carbon_callback_holder = callback_c

        # 2. Get the application event target. This is where we install the handler.
        _carbon.GetApplicationEventTarget.restype = ctypes.c_void_p
        target = _carbon.GetApplicationEventTarget()

        # 3. Install the event handler. Tell it to fire on hotkey-pressed events.
        event_type = _EventTypeSpec(
            eventClass=_KEVENT_CLASS_KEYBOARD,
            eventKind=_KEVENT_HOTKEY_PRESSED,
        )
        handler_ref = ctypes.c_void_p()
        _carbon.InstallEventHandler.argtypes = [
            ctypes.c_void_p,                  # EventTargetRef
            ctypes.c_void_p,                  # EventHandlerUPP (function ptr)
            ctypes.c_uint32,                  # numTypes
            ctypes.POINTER(_EventTypeSpec),   # typeList
            ctypes.c_void_p,                  # userData
            ctypes.POINTER(ctypes.c_void_p),  # outRef
        ]
        _carbon.InstallEventHandler.restype = ctypes.c_int32
        status = _carbon.InstallEventHandler(
            target,
            ctypes.cast(callback_c, ctypes.c_void_p),
            1,
            ctypes.byref(event_type),
            None,
            ctypes.byref(handler_ref),
        )
        if status != 0:
            sys.stderr.write(f"[tsifl-helper] InstallEventHandler failed (status={status})\n")
            return False
        _carbon_handler_ref = handler_ref

        # 4. Register the actual hotkey: Cmd+Option+T.
        hotkey_id = _EventHotKeyID(signature=0x74736966, id=1)  # 'tsif' as 4-char code
        hotkey_ref = ctypes.c_void_p()
        _carbon.RegisterEventHotKey.argtypes = [
            ctypes.c_uint32,                  # inHotKeyCode (virtual keycode)
            ctypes.c_uint32,                  # inHotKeyModifiers (cmd|shift|opt mask)
            _EventHotKeyID,                   # inHotKeyID
            ctypes.c_void_p,                  # inTarget
            ctypes.c_uint32,                  # inOptions
            ctypes.POINTER(ctypes.c_void_p),  # outRef
        ]
        _carbon.RegisterEventHotKey.restype = ctypes.c_int32
        status = _carbon.RegisterEventHotKey(
            _VK_T,
            _CMD_KEY | _OPTION_KEY,           # ⌘⌥ (avoids browser ⌘⇧T conflict)
            hotkey_id,
            target,
            0,
            ctypes.byref(hotkey_ref),
        )
        if status != 0:
            sys.stderr.write(f"[tsifl-helper] RegisterEventHotKey failed (status={status})\n")
            return False
        _carbon_hotkey_ref = hotkey_ref
        sys.stderr.write("[tsifl-helper] Global hotkey registered: ⌘⌥T (cmd+option+t)\n")
        return True
    except Exception as e:
        sys.stderr.write(f"[tsifl-helper] Carbon hotkey registration error: {e}\n")
        return False


# ── rumps menu bar app ──────────────────────────────────────────────────────

class TsiflHelperApp(rumps.App):
    """Menu bar app that surfaces agent status and provides controls.

    Icon convention:
      "tsifl"        → agent running, connected to backend
      "tsifl ●"      → agent running with errors (last_error set)
      "tsifl —"      → agent stopped / not running

    The menu lets the user view the log file, open the Anthropic console,
    or quit cleanly. No "start/stop" — the app is the agent.
    """

    def __init__(self):
        # Use a proper template icon for the menu bar. Template images are
        # black + transparent; macOS auto-tints based on light/dark mode.
        # Loaded from sibling `icon.png` (44x44 retina, generated by
        # _make_icon.py). Title left empty so only the icon shows.
        icon_path = self._resolve_icon_path()
        super().__init__(
            name="tsifl Helper",
            title="",
            icon=icon_path,
            template=True,        # macOS auto-tint for light/dark mode
            quit_button=None,
        )
        # Fallback icon-less mode for first launch — text-only menu bar entry.
        # Once we ship a proper .icns icon, we'll set self.icon = "icon.icns".
        self.menu = [
            rumps.MenuItem("Status: starting...", callback=None),
            None,  # separator
            rumps.MenuItem("Open with ⌘⌥T", callback=self.on_open_shortcut),
            rumps.MenuItem("Open Logs", callback=self.on_open_logs),
            rumps.MenuItem("Anthropic Console", callback=self.on_open_console),
            None,
            rumps.MenuItem("Quit tsifl Helper", callback=self.on_quit),
        ]

        # Kick off the agent right after the menu bar app is up
        _start_agent_thread()

        # Register the global Cmd+Shift+T hotkey. Failure is non-fatal —
        # the rest of the app still works, the user just doesn't get the
        # global shortcut until they grant Input Monitoring permission and
        # relaunch.
        global _app_ref
        _app_ref = self
        _hotkey_ok = _start_hotkey_listener()
        if not _hotkey_ok:
            sys.stderr.write(
                "[tsifl-helper] global ⌘⌥T listener could not attach via "
                "Carbon. The menu item 'Open with ⌘⌥T' still works. "
                "Carbon errors usually indicate a deeper macOS issue — "
                "check Console.app for security policy denials.\n"
            )

    @staticmethod
    def _resolve_icon_path() -> str | None:
        """Find icon.png whether running from source or from a bundled .app.

        - Dev: sibling of this file (desktop-agent/icon.png)
        - Bundled: in the .app's Resources directory (we add it via DATA_FILES
          in setup.py).

        Returns the absolute path string if found, else None (rumps falls
        back to the title text).
        """
        # Try sibling-of-this-file first (works for both dev and py2app
        # alias mode, since alias preserves the source paths)
        here_icon = Path(__file__).resolve().parent / "icon.png"
        if here_icon.exists():
            return str(here_icon)
        # Bundled .app standalone mode: icon ends up in Resources/
        if hasattr(sys, "_MEIPASS"):  # PyInstaller pattern; py2app uses different
            return str(Path(sys._MEIPASS) / "icon.png")
        # py2app standalone: walk up from sys.argv[0] to find Resources/
        try:
            launcher = Path(sys.argv[0]).resolve()
            for parent in launcher.parents:
                resources = parent / "Resources"
                if resources.exists() and (resources / "icon.png").exists():
                    return str(resources / "icon.png")
        except Exception:
            pass
        return None

        # Tick every 5s to update status
        self._status_timer = rumps.Timer(self.on_tick, 5)
        self._status_timer.start()

        # First-launch onboarding: shown once, marks itself complete so
        # subsequent launches skip it. Deferred 2s so the menu bar icon
        # is fully visible before the modal blocks the main thread.
        rumps.Timer(self._show_onboarding_once, 2).start()

    def _show_onboarding_once(self, timer):
        """Trigger the first-launch dialog exactly once, then stop the timer."""
        timer.stop()
        try:
            _maybe_show_first_launch_dialog()
        except Exception as e:
            sys.stderr.write(f"[tsifl-helper] onboarding dialog failed: {e}\n")

    # ── Menu callbacks ──────────────────────────────────────────────────────

    def on_open_shortcut(self, _):
        """Open the prompt panel from the menu (mirror of Cmd+Shift+T).
        Useful as a fallback when the global hotkey isn't working yet
        (Input Monitoring permission not granted, etc)."""
        try:
            _show_shortcut_panel()
        except Exception as e:
            rumps.alert(title="tsifl", message=f"Couldn't open panel: {e}")

    def on_open_logs(self, _):
        """Open the agent's rotating log file in the user's default app."""
        log_path = Path.home() / "Library" / "Logs" / "tsifl-agent.log"
        if log_path.exists():
            # 'open' is the macOS shell command; rumps doesn't expose a richer API
            import subprocess
            subprocess.Popen(["open", str(log_path)])
        else:
            rumps.alert(
                title="No log file yet",
                message=(
                    "The agent hasn't written any log entries yet. "
                    "If you just launched, give it a few seconds and try again."
                ),
            )

    def on_open_console(self, _):
        """Quick-access link to Anthropic billing / spend dashboard."""
        webbrowser.open("https://console.anthropic.com/settings/billing")

    def on_quit(self, _):
        """Confirm + exit cleanly."""
        ok = rumps.alert(
            title="Quit tsifl Helper?",
            message=(
                "Stopping the helper means tsifl can't run advanced Excel "
                "features (Solver, Goal Seek, Data Tables, SmartArt, etc.) "
                "until you launch it again."
            ),
            ok="Quit",
            cancel="Keep Running",
        )
        if ok:
            rumps.quit_application()

    # ── Status tick ─────────────────────────────────────────────────────────

    def on_tick(self, _):
        """Update the menu bar title + status menu item every 5s."""
        # Find the "Status" menu item (always at index 0)
        status_item = self.menu.get("Status: starting...") or next(iter(self.menu.values()))
        # Title stays empty in normal state (icon-only). On error, append a
        # bullet-mark to the icon so the user knows something's wrong without
        # opening the menu. On stopped, append a dash. Icon itself doesn't
        # change — rumps would need a full image swap which is heavier.
        if _last_error:
            self.title = "  ●"  # leading spaces for spacing from icon
            status_text = f"Error: {_last_error[:40]}..."
        elif _agent_thread is None or not _agent_thread.is_alive():
            self.title = "  —"
            status_text = "Status: stopped"
        else:
            self.title = ""
            uptime = int(time.time() - (_agent_started_at or time.time()))
            mins = uptime // 60
            secs = uptime % 60
            if mins >= 60:
                hrs = mins // 60
                mins = mins % 60
                status_text = f"Status: connected · up {hrs}h {mins}m"
            else:
                status_text = f"Status: connected · up {mins}m {secs}s"

        # rumps doesn't expose a clean way to rename a MenuItem dynamically,
        # so we re-build the menu with the new status text in slot 0.
        # (The other items are static so this is cheap.)
        try:
            # MenuItem objects are keyed by their initial title — find by reference
            for key in list(self.menu.keys()):
                if key.startswith("Status:"):
                    del self.menu[key]
                    break
            new_item = rumps.MenuItem(status_text, callback=None)
            self.menu.insert_after(None, new_item) if False else None
            # rumps doesn't have insert-at-index, so we re-add as the first item
            # by clearing and rebuilding. Actually simplest: just always show
            # status as the menu's first key. Set it on the app:
            self.menu = [
                new_item,
                None,
                rumps.MenuItem("Open Logs", callback=self.on_open_logs),
                rumps.MenuItem("Anthropic Console", callback=self.on_open_console),
                None,
                rumps.MenuItem("Quit tsifl Helper", callback=self.on_quit),
            ]
        except Exception:
            # Menu rebuild failed (rumps internals shifted) — just update title
            pass


def main():
    """Entry point for both `python3 tsifl_helper_app.py` and the bundled .app."""
    app = TsiflHelperApp()
    app.run()


if __name__ == "__main__":
    main()
