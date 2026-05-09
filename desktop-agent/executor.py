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
    """Launch or activate a macOS application.

    Uses `open -a` which returns immediately (non-blocking) — avoids
    the 15s AppleScript timeout on cold launches.  The vision loop's
    `wait` action gives the app time to fully appear before screenshotting.
    """
    try:
        subprocess.run(["open", "-a", app_name], check=True, timeout=10)
        # Give the app a moment to register, then try to bring it front
        time.sleep(0.5)
        # Quick activate — if the app is already running this is instant;
        # if it's still loading we don't care if this times out (5s cap)
        run_applescript(
            f'tell application "{app_name}" to activate', timeout=5
        )
        return True, f"Opened {app_name}"
    except subprocess.CalledProcessError:
        return False, f"Could not find app: {app_name}"
    except Exception as e:
        # Even if activate timed out, open -a likely succeeded
        return True, f"Opened {app_name} (still loading)"


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


# ── Screen automation (vision loop primitives) ──────────────────────────────
# These let Claude see the screen and interact with ANY app — no per-app API
# needed. The flow: screenshot → Claude analyzes → click/type/scroll → repeat.

def _get_screen_point_size() -> tuple[int, int]:
    """Get the main display size in screen POINTS (not retina pixels).

    On a retina Mac the physical pixels are 2× the point dimensions.
    CGEvent clicks use points, so we need this for coordinate mapping.
    """
    try:
        import Quartz
        main = Quartz.CGMainDisplayID()
        w = int(Quartz.CGDisplayPixelsWide(main))   # point width
        h = int(Quartz.CGDisplayPixelsHigh(main))    # point height
        return w, h
    except Exception:
        # Fallback — query via system_profiler
        try:
            out = subprocess.run(
                ["system_profiler", "SPDisplaysDataType"],
                capture_output=True, text=True, timeout=5,
            ).stdout
            import re
            m = re.search(r"Resolution:\s*(\d+)\s*x\s*(\d+)", out)
            if m:
                return int(m.group(1)), int(m.group(2))
        except Exception:
            pass
        return 1470, 956  # sensible MacBook default


def capture_screenshot(region: dict = None) -> tuple[bool, str]:
    """Capture the screen and return base64-encoded JPEG.

    CRITICAL — Retina coordinate mapping:
      screencapture produces 2× retina pixels (e.g. 2940×1912), but
      CGEvent clicks use screen POINTS (e.g. 1470×956).  We resize the
      screenshot to the screen's point dimensions so that when Claude
      says "click at (x, y)" those coordinates map 1:1 to screen points.

    Args:
        region: optional {x, y, width, height} to capture a sub-region
                (in screen points).  If None, captures the full display.

    Returns:
        (True, base64_jpeg_string) on success.
    """
    import base64
    import tempfile

    raw_path = os.path.join(tempfile.gettempdir(), "tsifl_screen.png")
    try:
        cmd = ["screencapture", "-x"]  # -x = no sound
        if region:
            x, y = region.get("x", 0), region.get("y", 0)
            w, h = region.get("width", 800), region.get("height", 600)
            cmd.extend(["-R", f"{x},{y},{w},{h}"])
        cmd.append(raw_path)
        subprocess.run(cmd, timeout=5, check=True)

        # Resize to screen POINT dimensions so coordinates map 1:1 to clicks
        try:
            from PIL import Image as _PILImage
            import io as _io
            img = _PILImage.open(raw_path)

            # Convert RGBA→RGB (JPEG can't do alpha)
            if img.mode in ("RGBA", "LA"):
                bg = _PILImage.new("RGB", img.size, (255, 255, 255))
                bg.paste(img, mask=img.split()[-1])
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")

            # Resize to screen point dimensions (not arbitrary max)
            screen_w, screen_h = _get_screen_point_size()
            raw_w, raw_h = img.size
            logger.info(
                f"Screenshot: raw {raw_w}×{raw_h}, "
                f"screen points {screen_w}×{screen_h}, "
                f"scale {raw_w / screen_w:.1f}×"
            )

            if raw_w != screen_w or raw_h != screen_h:
                img = img.resize((screen_w, screen_h), _PILImage.LANCZOS)

            # Encode as JPEG — shrink until under 700KB
            for quality in (80, 65, 50, 40):
                buf = _io.BytesIO()
                img.save(buf, format="JPEG", quality=quality, optimize=True)
                data = buf.getvalue()
                if len(data) <= 700_000:
                    break
            return True, base64.b64encode(data).decode("ascii")
        except ImportError:
            # PIL not available — send raw PNG (larger but still works)
            with open(raw_path, "rb") as f:
                data = base64.b64encode(f.read()).decode("ascii")
            return True, data
    except Exception as e:
        return False, f"Screenshot failed: {e}"


def click_at_position(x: int, y: int, click_type: str = "left") -> tuple[bool, str]:
    """Click at screen coordinates using CGEvent (Quartz).

    Works with any app — no Accessibility API needed for basic clicks.

    Args:
        x, y: screen coordinates (pixels from top-left)
        click_type: "left", "right", or "double"
    """
    try:
        import Quartz

        point = (float(x), float(y))

        if click_type == "right":
            down_type = Quartz.kCGEventRightMouseDown
            up_type = Quartz.kCGEventRightMouseUp
            button = Quartz.kCGMouseButtonRight
        else:
            down_type = Quartz.kCGEventLeftMouseDown
            up_type = Quartz.kCGEventLeftMouseUp
            button = Quartz.kCGMouseButtonLeft

        # Move mouse first (some apps need this)
        move = Quartz.CGEventCreateMouseEvent(
            None, Quartz.kCGEventMouseMoved, point, button
        )
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, move)
        time.sleep(0.05)

        # Click down
        down = Quartz.CGEventCreateMouseEvent(None, down_type, point, button)
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, down)
        time.sleep(0.05)

        # Click up
        up = Quartz.CGEventCreateMouseEvent(None, up_type, point, button)
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, up)

        if click_type == "double":
            time.sleep(0.1)
            down2 = Quartz.CGEventCreateMouseEvent(None, down_type, point, button)
            Quartz.CGEventSetIntegerValueField(
                down2, Quartz.kCGMouseEventClickState, 2
            )
            Quartz.CGEventPost(Quartz.kCGHIDEventTap, down2)
            time.sleep(0.05)
            up2 = Quartz.CGEventCreateMouseEvent(None, up_type, point, button)
            Quartz.CGEventSetIntegerValueField(
                up2, Quartz.kCGMouseEventClickState, 2
            )
            Quartz.CGEventPost(Quartz.kCGHIDEventTap, up2)

        return True, f"Clicked at ({x}, {y})"
    except ImportError:
        # Fallback: AppleScript — less precise but works
        return run_applescript(
            f'tell application "System Events" to click at {{{x}, {y}}}'
        )
    except Exception as e:
        return False, f"Click failed: {e}"


def type_text_keyboard(text: str) -> tuple[bool, str]:
    """Type text using System Events keystrokes.

    Handles special characters by using 'keystroke' for printable chars.
    For large blocks of text, uses clipboard paste for speed.
    """
    if not text:
        return True, "Nothing to type"

    # For long text (>50 chars), paste from clipboard is faster + more reliable
    if len(text) > 50:
        ok, _ = set_clipboard(text)
        if ok:
            return run_applescript(
                'tell application "System Events" to keystroke "v" using command down'
            )

    # Short text: type directly (preserves clipboard)
    # Escape backslashes and quotes for AppleScript
    escaped = text.replace("\\", "\\\\").replace('"', '\\"')
    return run_applescript(
        f'tell application "System Events" to keystroke "{escaped}"'
    )


def press_key_combo(keys: str) -> tuple[bool, str]:
    """Press a keyboard shortcut.

    Args:
        keys: shortcut string like "cmd+c", "cmd+shift+t", "return",
              "tab", "escape", "space", "delete", "up", "down"

    Supports: cmd, shift, ctrl/control, alt/option + any key.
    """
    parts = [k.strip().lower() for k in keys.split("+")]

    # Map modifier names to AppleScript modifiers
    modifier_map = {
        "cmd": "command down",
        "command": "command down",
        "shift": "shift down",
        "ctrl": "control down",
        "control": "control down",
        "alt": "option down",
        "option": "option down",
    }

    # Map special key names to AppleScript key codes
    special_keys = {
        "return": 36, "enter": 36,
        "tab": 48,
        "escape": 53, "esc": 53,
        "space": 49,
        "delete": 51, "backspace": 51,
        "forwarddelete": 117,
        "up": 126, "down": 125, "left": 123, "right": 124,
        "home": 115, "end": 119,
        "pageup": 116, "pagedown": 121,
        "f1": 122, "f2": 120, "f3": 99, "f4": 118,
        "f5": 96, "f6": 97, "f7": 98, "f8": 100,
    }

    modifiers = []
    key_part = None

    for p in parts:
        if p in modifier_map:
            modifiers.append(modifier_map[p])
        else:
            key_part = p

    if key_part is None:
        return False, f"No key specified in '{keys}'"

    modifier_str = ", ".join(modifiers)

    if key_part in special_keys:
        # Use key code for special keys
        code = special_keys[key_part]
        if modifiers:
            script = f'tell application "System Events" to key code {code} using {{{modifier_str}}}'
        else:
            script = f'tell application "System Events" to key code {code}'
    else:
        # Regular character
        escaped = key_part.replace("\\", "\\\\").replace('"', '\\"')
        if modifiers:
            script = f'tell application "System Events" to keystroke "{escaped}" using {{{modifier_str}}}'
        else:
            script = f'tell application "System Events" to keystroke "{escaped}"'

    return run_applescript(script)


def scroll_screen(direction: str = "down", amount: int = 3, x: int = 0, y: int = 0) -> tuple[bool, str]:
    """Scroll at the current mouse position or specified coordinates.

    Args:
        direction: "up" or "down"
        amount: number of scroll ticks (3 ≈ one page-ish)
        x, y: optional coordinates to scroll at (0,0 = current position)
    """
    try:
        import Quartz

        scroll_amount = amount if direction == "up" else -amount

        # Move mouse to position first if coordinates given
        if x > 0 and y > 0:
            move = Quartz.CGEventCreateMouseEvent(
                None, Quartz.kCGEventMouseMoved,
                (float(x), float(y)), Quartz.kCGMouseButtonLeft,
            )
            Quartz.CGEventPost(Quartz.kCGHIDEventTap, move)
            time.sleep(0.05)

        event = Quartz.CGEventCreateScrollWheelEvent(
            None, Quartz.kCGScrollEventUnitLine, 1, scroll_amount
        )
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)
        return True, f"Scrolled {direction} {amount}"
    except ImportError:
        # Fallback to AppleScript
        key_code = 126 if direction == "up" else 125  # arrow keys
        script = f'tell application "System Events" to key code {key_code}'
        for _ in range(amount):
            run_applescript(script)
        return True, f"Scrolled {direction} {amount}"
    except Exception as e:
        return False, f"Scroll failed: {e}"


def wait_seconds(seconds: float) -> tuple[bool, str]:
    """Wait for a specified duration (for UI to settle)."""
    seconds = min(seconds, 10)  # cap at 10s
    time.sleep(seconds)
    return True, f"Waited {seconds}s"


# ── Spotify (instant play via AppleScript — no vision loop needed) ────────

def spotify_play(query: str) -> tuple[bool, str]:
    """Play on Spotify — pure AppleScript keystrokes, no CGEvent focus issues.

    Flow (single AppleScript block keeps Spotify focused throughout):
    1. Activate Spotify + set frontmost
    2. Cmd+K → opens search overlay
    3. Cmd+A → select existing text, then type query
    4. Wait for suggestions → Down arrow → Enter to play top result
    5. Verify playback, fallback to Enter/Space if needed

    Total time: ~8 seconds. No CGEvent = no focus-stealing.
    """
    escaped = query.replace("\\", "\\\\").replace('"', '\\"')

    try:
        # Ensure Spotify is running (non-blocking launch)
        subprocess.run(["open", "-a", "Spotify"], check=True, timeout=10)
        time.sleep(1)

        # All-in-one AppleScript — Spotify keeps focus the entire time.
        # Cmd+K opens search, type query, Down picks first suggestion, Enter plays.
        ok, result = run_applescript(
            'tell application "Spotify" to activate\n'
            'delay 0.8\n'
            'tell application "System Events" to tell process "Spotify"\n'
            '    set frontmost to true\n'
            '    delay 0.3\n'
            '    keystroke "k" using command down\n'
            '    delay 0.5\n'
            '    keystroke "a" using command down\n'
            '    delay 0.1\n'
            f'    keystroke "{escaped}"\n'
            '    delay 2\n'
            '    key code 125\n'
            '    delay 0.3\n'
            '    key code 36\n'
            'end tell',
            timeout=15,
        )
        time.sleep(2)

        # ── Verify playback ──────────────────────────────────────────
        def _check_playing() -> tuple[bool, str]:
            ok_s, info = run_applescript(
                'tell application "Spotify"\n'
                '    if player state is playing then\n'
                '        return "▶️ " & name of current track & " — " & artist of current track\n'
                '    else\n'
                '        return "NOT_PLAYING"\n'
                '    end if\n'
                'end tell',
                timeout=5,
            )
            return ok_s, info or "NOT_PLAYING"

        ok_state, info = _check_playing()

        # Fallback 1: press Enter again (may have landed on artist page)
        if "NOT_PLAYING" in info:
            logger.info("spotify_play: first attempt didn't play — pressing Enter again")
            run_applescript(
                'tell application "System Events" to tell process "Spotify"\n'
                '    set frontmost to true\n'
                '    delay 0.3\n'
                '    key code 36\n'
                'end tell',
                timeout=5,
            )
            time.sleep(2)
            ok_state, info = _check_playing()

        # Fallback 2: Space bar (universal play/pause toggle)
        if "NOT_PLAYING" in info:
            logger.info("spotify_play: trying Space as last resort")
            run_applescript(
                'tell application "System Events" to tell process "Spotify"\n'
                '    set frontmost to true\n'
                '    delay 0.3\n'
                '    key code 49\n'
                'end tell',
                timeout=5,
            )
            time.sleep(1.5)
            ok_state, info = _check_playing()

        if "NOT_PLAYING" in info:
            info = f"Searched for {query} but playback didn't start"

        return ok_state, info

    except Exception as e:
        return False, f"Spotify play failed: {e}"


# ── Gmail operations (local Gmail API via OAuth token) ──────────────────────
# The desktop agent talks to Gmail directly using the local OAuth token at
# ~/.tsifulator_gmail_token.json. No backend round-trip needed — faster and
# works offline (for reads). If the token doesn't exist, we tell the user
# to run the setup script.

_GMAIL_TOKEN_PATH = Path.home() / ".tsifulator_gmail_token.json"


def _get_gmail_service():
    """Build an authenticated Gmail API service from the local token.

    Returns the service object, or raises RuntimeError with a user-friendly
    message if setup is needed.
    """
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except ImportError:
        raise RuntimeError(
            "Gmail libraries missing. Run:\n"
            "  pip install google-api-python-client google-auth-oauthlib"
        )

    if not _GMAIL_TOKEN_PATH.exists():
        raise RuntimeError(
            "Gmail not connected yet. Run:\n"
            "  cd tsifulator.ai && python3 gmail-client/gmail_setup.py"
        )

    creds = Credentials.from_authorized_user_file(str(_GMAIL_TOKEN_PATH))

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        _GMAIL_TOKEN_PATH.write_text(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def _extract_email_body(payload: dict) -> str:
    """Recursively extract plain text body from Gmail message payload."""
    import base64
    if payload.get("mimeType") == "text/plain":
        data = payload.get("body", {}).get("data", "")
        if data:
            return base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="replace")
    if "parts" in payload:
        for part in payload["parts"]:
            result = _extract_email_body(part)
            if result:
                return result
    return ""


def gmail_check_inbox(max_results: int = 10) -> tuple[bool, str]:
    """Fetch recent inbox messages using the local Gmail API."""
    try:
        service = _get_gmail_service()
        results = service.users().messages().list(
            userId="me", labelIds=["INBOX"], maxResults=max_results
        ).execute()

        messages = results.get("messages", [])
        if not messages:
            return True, "Inbox is empty."

        lines = [f"📬 {len(messages)} recent emails:"]
        for i, msg in enumerate(messages, 1):
            detail = service.users().messages().get(
                userId="me", id=msg["id"],
                format="metadata",
                metadataHeaders=["From", "Subject", "Date"],
            ).execute()
            headers = {h["name"]: h["value"] for h in detail["payload"]["headers"]}
            unread = "📩" if "UNREAD" in detail.get("labelIds", []) else "  "
            sender = headers.get("From", "?")
            if "<" in sender:
                sender = sender.split("<")[0].strip().strip('"')
            lines.append(f"{unread} {i}. {sender}")
            lines.append(f"     {headers.get('Subject', '(no subject)')}")
            snippet = detail.get("snippet", "")
            if snippet:
                lines.append(f"     {snippet[:100]}")
            lines.append(f"     [id:{msg['id']}]")
        return True, "\n".join(lines)
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail error: {e}"


def gmail_search(query: str, max_results: int = 10) -> tuple[bool, str]:
    """Search emails using Gmail query syntax."""
    try:
        service = _get_gmail_service()
        results = service.users().messages().list(
            userId="me", q=query, maxResults=max_results
        ).execute()

        messages = results.get("messages", [])
        if not messages:
            return True, f"No emails found for '{query}'."

        lines = [f"Found {len(messages)} email(s) for '{query}':"]
        for i, msg in enumerate(messages, 1):
            detail = service.users().messages().get(
                userId="me", id=msg["id"],
                format="metadata",
                metadataHeaders=["From", "Subject", "Date"],
            ).execute()
            headers = {h["name"]: h["value"] for h in detail["payload"]["headers"]}
            sender = headers.get("From", "?")
            if "<" in sender:
                sender = sender.split("<")[0].strip().strip('"')
            lines.append(f"  {i}. {sender}")
            lines.append(f"     {headers.get('Subject', '(no subject)')}")
            snippet = detail.get("snippet", "")
            if snippet:
                lines.append(f"     {snippet[:100]}")
            lines.append(f"     [id:{msg['id']}]")
        return True, "\n".join(lines)
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail search failed: {e}"


def gmail_read_message(message_id: str) -> tuple[bool, str]:
    """Read the full body of a specific email message."""
    try:
        service = _get_gmail_service()
        msg = service.users().messages().get(
            userId="me", id=message_id, format="full"
        ).execute()

        headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
        body_text = _extract_email_body(msg["payload"])
        if len(body_text) > 2000:
            body_text = body_text[:1997] + "…"

        lines = [
            f"From: {headers.get('From', '?')}",
            f"To: {headers.get('To', '?')}",
            f"Subject: {headers.get('Subject', '(no subject)')}",
            f"Date: {headers.get('Date', '?')}",
            "",
            body_text or "(no body)",
        ]
        return True, "\n".join(lines)
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail read failed: {e}"


def gmail_send(to: str, subject: str, body: str, reply_to_id: str = "") -> tuple[bool, str]:
    """Send an email (or reply to a thread) via Gmail API."""
    try:
        import base64
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart

        service = _get_gmail_service()
        msg = MIMEMultipart("alternative")
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        send_body = {"raw": raw}
        if reply_to_id:
            send_body["threadId"] = reply_to_id

        sent = service.users().messages().send(userId="me", body=send_body).execute()
        return True, f"✉️ Email sent to {to}"
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail send failed: {e}"


def gmail_draft(to: str, subject: str, body: str) -> tuple[bool, str]:
    """Create a draft email in Gmail (doesn't send)."""
    try:
        import base64
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart

        service = _get_gmail_service()
        msg = MIMEMultipart("alternative")
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        draft = service.users().drafts().create(
            userId="me", body={"message": {"raw": raw}}
        ).execute()
        return True, f"📝 Draft created for {to}"
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail draft failed: {e}"


# ── Action executor ──────────────────────────────────────────────────────────

def execute_action(action: Action) -> Action:
    """Execute a single action and update it with the result.

    Returns the same Action with `result`, `success`, and `error` filled in.
    """
    try:
        # Parse command — might be a raw string or a JSON payload dict
        cmd = action.command
        cmd_data = {}
        if cmd.startswith("{"):
            try:
                cmd_data = json.loads(cmd)
            except json.JSONDecodeError:
                pass

        if action.type == "search_files":
            params = cmd_data if cmd_data else {"query": cmd}
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
            path = cmd_data.get("path", cmd_data.get("file_path", cmd))
            action.success, action.result = open_file(path)

        elif action.type == "open_app":
            app_name = cmd_data.get("app", cmd_data.get("app_name", cmd_data.get("name", cmd)))
            action.success, action.result = open_app(app_name)

        elif action.type == "open_url":
            url = cmd_data.get("url", cmd)
            action.success, action.result = open_url(url)

        elif action.type == "applescript":
            script = cmd_data.get("script", cmd_data.get("code", cmd))
            action.success, action.result = run_applescript(script)

        elif action.type == "shell":
            shell_cmd = cmd_data.get("command", cmd)
            action.success, action.result = run_shell(shell_cmd)

        elif action.type == "clipboard_copy":
            text = cmd_data.get("text", cmd)
            action.success, action.result = set_clipboard(text)

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

        # ── Screen automation actions ───────────────────────────────────
        elif action.type == "screenshot":
            region = cmd_data.get("region") if cmd_data else None
            action.success, action.result = capture_screenshot(region)

        elif action.type == "click_at":
            x = cmd_data.get("x", 0)
            y = cmd_data.get("y", 0)
            ct = cmd_data.get("click_type", "left")
            action.success, action.result = click_at_position(x, y, ct)

        elif action.type == "type_text":
            text = cmd_data.get("text", cmd)
            action.success, action.result = type_text_keyboard(text)

        elif action.type == "key_combo":
            keys = cmd_data.get("keys", cmd)
            action.success, action.result = press_key_combo(keys)

        elif action.type == "scroll":
            direction = cmd_data.get("direction", "down")
            amount = cmd_data.get("amount", 3)
            sx = cmd_data.get("x", 0)
            sy = cmd_data.get("y", 0)
            action.success, action.result = scroll_screen(direction, amount, sx, sy)

        elif action.type == "wait":
            secs = cmd_data.get("seconds", 1)
            action.success, action.result = wait_seconds(secs)

        # ── Gmail actions ───────────────────────────────────────────────
        elif action.type == "check_inbox":
            max_r = cmd_data.get("max_results", 10)
            action.success, action.result = gmail_check_inbox(max_r)

        elif action.type == "search_email":
            query = cmd_data.get("query", cmd)
            max_r = cmd_data.get("max_results", 10)
            action.success, action.result = gmail_search(query, max_r)

        elif action.type == "read_email":
            msg_id = cmd_data.get("message_id", cmd)
            action.success, action.result = gmail_read_message(msg_id)

        elif action.type == "send_email":
            action.success, action.result = gmail_send(
                to=cmd_data.get("to", ""),
                subject=cmd_data.get("subject", ""),
                body=cmd_data.get("body", ""),
                reply_to_id=cmd_data.get("reply_to_id", ""),
            )

        elif action.type == "draft_email":
            action.success, action.result = gmail_draft(
                to=cmd_data.get("to", ""),
                subject=cmd_data.get("subject", ""),
                body=cmd_data.get("body", ""),
            )

        elif action.type == "spotify_play":
            query = cmd_data.get("query", cmd)
            action.success, action.result = spotify_play(query)

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
