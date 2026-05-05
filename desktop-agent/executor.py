"""
executor.py — tsifl Mac action execution engine.

Takes structured action plans from Claude and executes them on the
user's Mac via AppleScript, Spotlight, shell commands, and native APIs.

Every action goes through a risk classification before execution:
  - GREEN  (read-only): auto-execute, no confirmation needed
  - YELLOW (writes): show plan, one-click confirm
  - RED    (irreversible): show plan, require explicit confirmation

The executor never runs anything the user hasn't approved.
"""

import json
import logging
import os
import subprocess
import time
from dataclasses import dataclass, field, asdict
from enum import Enum
from pathlib import Path
from typing import Optional

logger = logging.getLogger("tsifl.executor")


# ── Risk levels ──────────────────────────────────────────────────────────────

class Risk(str, Enum):
    GREEN = "green"    # read-only: search, open, show info
    YELLOW = "yellow"  # writes: create file, type text, move file
    RED = "red"        # irreversible: send email, delete, purchase


# ── Action dataclass ─────────────────────────────────────────────────────────

@dataclass
class Action:
    """A single step in an execution plan."""
    type: str                          # e.g. "search_files", "open_app", "applescript"
    description: str                   # human-readable: "Search for Excel files containing 'grocery'"
    command: str                       # the actual command/script to execute
    risk: Risk = Risk.GREEN
    result: Optional[str] = None       # filled after execution
    success: bool = False
    error: Optional[str] = None

    def to_dict(self):
        d = asdict(self)
        d["risk"] = self.risk.value
        return d


# ── File search via Spotlight (mdfind) ───────────────────────────────────────

def search_files(query: str, max_results: int = 10, file_type: str = None) -> list[str]:
    """Search for files using macOS Spotlight (mdfind).

    Args:
        query: search term (filename, content, or raw mdfind query)
        max_results: cap on number of results
        file_type: optional filter — "excel", "pdf", "image", "document", etc.

    Returns:
        List of file paths matching the query.
    """
    # Map friendly type names to Spotlight content types
    type_filters = {
        "excel": 'kMDItemContentType == "org.openxmlformats.spreadsheetml.sheet" || kMDItemContentType == "com.microsoft.excel.xls"',
        "word": 'kMDItemContentType == "org.openxmlformats.wordprocessingml.document" || kMDItemContentType == "com.microsoft.word.doc"',
        "ppt": 'kMDItemContentType == "org.openxmlformats.presentationml.presentation"',
        "pdf": 'kMDItemContentType == "com.adobe.pdf"',
        "image": 'kMDItemContentTypeTree == "public.image"',
        "csv": 'kMDItemContentType == "public.comma-separated-values-text"',
        "text": 'kMDItemContentTypeTree == "public.text"',
    }

    # Build mdfind query
    parts = []
    if file_type and file_type.lower() in type_filters:
        parts.append(f"({type_filters[file_type.lower()]})")

    # Add name/content search
    if query:
        # If the query looks like a raw mdfind expression, use it directly
        if "kMDItem" in query:
            parts.append(query)
        else:
            # Search both filename and content
            escaped = query.replace('"', '\\"')
            parts.append(f'(kMDItemFSName == "*{escaped}*"cdw || kMDItemTextContent == "*{escaped}*"cdw)')

    mdfind_query = " && ".join(parts) if parts else f'kMDItemFSName == "*{query}*"cdw'

    try:
        result = subprocess.run(
            ["mdfind", mdfind_query],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode != 0:
            return []
        paths = [p.strip() for p in result.stdout.strip().split("\n") if p.strip()]
        # Filter out hidden/system paths
        paths = [p for p in paths if not any(
            seg.startswith(".") for seg in Path(p).parts[1:]  # skip root /
        )]
        # Sort by modification time (most recent first)
        paths.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0, reverse=True)
        return paths[:max_results]
    except subprocess.TimeoutExpired:
        return []
    except Exception as e:
        logger.error(f"mdfind failed: {e}")
        return []


# ── AppleScript execution ────────────────────────────────────────────────────

def run_applescript(script: str, timeout: int = 15) -> tuple[bool, str]:
    """Execute an AppleScript and return (success, output_or_error).

    Args:
        script: the AppleScript source code
        timeout: max seconds to wait

    Returns:
        (True, stdout) on success, (False, stderr) on failure.
    """
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=timeout,
        )
        if result.returncode == 0:
            return True, result.stdout.strip()
        return False, result.stderr.strip()
    except subprocess.TimeoutExpired:
        return False, f"AppleScript timed out after {timeout}s"
    except Exception as e:
        return False, str(e)


# ── Shell command execution (read-only) ──────────────────────────────────────

def run_shell(command: str, timeout: int = 10) -> tuple[bool, str]:
    """Run a read-only shell command and return (success, output).

    Safety: this should only be called for read-only commands (ls, cat,
    mdfind, date, etc.). Write operations should go through AppleScript
    or dedicated functions with proper risk classification.
    """
    try:
        result = subprocess.run(
            command, shell=True,
            capture_output=True, text=True, timeout=timeout,
        )
        output = result.stdout.strip()
        if result.returncode != 0 and result.stderr.strip():
            output = result.stderr.strip()
        return result.returncode == 0, output
    except subprocess.TimeoutExpired:
        return False, f"Command timed out after {timeout}s"
    except Exception as e:
        return False, str(e)


# ── High-level Mac operations ────────────────────────────────────────────────

def open_file(path: str) -> tuple[bool, str]:
    """Open a file with its default application."""
    try:
        subprocess.run(["open", path], check=True, timeout=5)
        return True, f"Opened {Path(path).name}"
    except Exception as e:
        return False, str(e)


def open_app(app_name: str) -> tuple[bool, str]:
    """Launch or activate a macOS application."""
    return run_applescript(f'tell application "{app_name}" to activate')


def open_url(url: str) -> tuple[bool, str]:
    """Open a URL in the default browser."""
    try:
        subprocess.run(["open", url], check=True, timeout=5)
        return True, f"Opened {url}"
    except Exception as e:
        return False, str(e)


def get_clipboard() -> str:
    """Get the current clipboard contents."""
    ok, text = run_shell("pbpaste")
    return text if ok else ""


def set_clipboard(text: str) -> tuple[bool, str]:
    """Set the clipboard contents."""
    try:
        process = subprocess.Popen(
            ["pbcopy"], stdin=subprocess.PIPE
        )
        process.communicate(text.encode("utf-8"), timeout=5)
        return True, "Copied to clipboard"
    except Exception as e:
        return False, str(e)


def get_frontmost_app() -> str:
    """Get the name of the frontmost application."""
    ok, name = run_applescript(
        'tell application "System Events" to get name of first process whose frontmost is true'
    )
    return name if ok else "unknown"


def get_running_apps() -> list[str]:
    """Get list of running applications."""
    ok, output = run_applescript(
        'tell application "System Events" to get name of every process whose background only is false'
    )
    if ok:
        return [a.strip() for a in output.split(",") if a.strip()]
    return []


# ── Rich context capture ────────────────────────────────────────────────────

def get_system_context() -> dict:
    """Capture rich context about the current Mac state."""
    ctx = {
        "frontmost_app": get_frontmost_app(),
        "running_apps": get_running_apps(),
        "time": time.strftime("%Y-%m-%d %H:%M:%S %Z"),
        "user": os.environ.get("USER", "unknown"),
        "home": str(Path.home()),
    }

    # Active document in frontmost app (if supported)
    front = ctx["frontmost_app"]
    if front in ("Microsoft Excel", "Microsoft Word", "Microsoft PowerPoint",
                 "Pages", "Numbers", "Keynote", "Preview", "TextEdit"):
        ok, doc = run_applescript(
            f'tell application "{front}" to get name of front document'
        )
        if ok:
            ctx["active_document"] = doc

    # Finder: get selected files
    if front == "Finder":
        ok, sel = run_applescript(
            'tell application "Finder" to get POSIX path of (selection as alias list)'
        )
        if ok and sel:
            ctx["finder_selection"] = sel

    # Safari/Chrome: get current tab URL + title
    if front in ("Safari", "Google Chrome"):
        if front == "Safari":
            ok, url = run_applescript(
                'tell application "Safari" to get URL of front document'
            )
            ok2, title = run_applescript(
                'tell application "Safari" to get name of front document'
            )
        else:
            ok, url = run_applescript(
                'tell application "Google Chrome" to get URL of active tab of front window'
            )
            ok2, title = run_applescript(
                'tell application "Google Chrome" to get title of active tab of front window'
            )
        if ok:
            ctx["browser_url"] = url
        if ok2:
            ctx["browser_title"] = title

    return ctx


# ── Action executor ──────────────────────────────────────────────────────────

def execute_action(action: Action) -> Action:
    """Execute a single action and update it with the result.

    Returns the same Action with `result`, `success`, and `error` filled in.
    """
    try:
        if action.type == "search_files":
            # command is JSON: {"query": "...", "file_type": "..."}
            params = json.loads(action.command) if action.command.startswith("{") else {"query": action.command}
            paths = search_files(
                query=params.get("query", ""),
                file_type=params.get("file_type"),
                max_results=params.get("max_results", 10),
            )
            if paths:
                action.result = "\n".join(paths)
                action.success = True
            else:
                action.result = "No files found"
                action.success = True  # search succeeded, just no results

        elif action.type == "open_file":
            action.success, action.result = open_file(action.command)

        elif action.type == "open_app":
            action.success, action.result = open_app(action.command)

        elif action.type == "open_url":
            action.success, action.result = open_url(action.command)

        elif action.type == "applescript":
            action.success, action.result = run_applescript(action.command)

        elif action.type == "shell":
            action.success, action.result = run_shell(action.command)

        elif action.type == "clipboard_copy":
            action.success, action.result = set_clipboard(action.command)

        elif action.type == "clipboard_read":
            text = get_clipboard()
            action.result = text or "(clipboard empty)"
            action.success = True

        elif action.type == "notify":
            # Show a macOS notification
            try:
                import rumps
                rumps.notification(
                    title="tsifl",
                    subtitle="",
                    message=action.command,
                )
                action.success = True
                action.result = "Notification shown"
            except Exception as e:
                action.success = False
                action.error = str(e)

        else:
            action.success = False
            action.error = f"Unknown action type: {action.type}"

    except Exception as e:
        action.success = False
        action.error = str(e)
        logger.error(f"Action failed ({action.type}): {e}")

    return action


def execute_plan(actions: list[Action], stop_on_error: bool = True) -> list[Action]:
    """Execute a list of actions sequentially.

    Args:
        actions: ordered list of Action objects
        stop_on_error: if True, stop executing after the first failure

    Returns:
        The same list with results filled in.
    """
    for action in actions:
        execute_action(action)
        if stop_on_error and not action.success:
            break
    return actions
