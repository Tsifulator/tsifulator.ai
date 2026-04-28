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
    app_map = {
        "Microsoft Excel": "excel",
        "RStudio": "rstudio",
        "Microsoft PowerPoint": "powerpoint",
        "Microsoft Word": "word",
    }
    app = app_map.get(frontmost_app, "shortcut")

    body = {
        "user_id": "shortcut-anon",
        "message": message,
        "context": {
            "app": app,
            "source": "global-shortcut",
            "frontmost_app": frontmost_app,
        },
    }
    if images:
        body["images"] = images

    try:
        # Image-attached requests can be larger so we bump the timeout
        timeout = 120 if images else 60
        with _httpx.Client(timeout=timeout) as client:
            r = client.post(f"{backend.rstrip('/')}/chat/", json=body)
            if r.status_code == 200:
                return r.json()
            return {"reply": f"Backend error ({r.status_code}): {r.text[:200]}",
                    "actions": []}
    except Exception as e:
        return {"reply": f"Could not reach tsifl: {e}", "actions": []}


# ── Floating shortcut panel (Spotlight-style) ───────────────────────────────
# Custom NSPanel built via PyObjC. Replaces the previous rumps.Window modal.
# Design specs (from user):
#   - Petite, "almost like its just the box alone"
#   - Clean white background with tsifl-blue outline
#   - Small `t` logo on the left
#   - Floats above all apps, joins all spaces (Mission Control compatible)
#   - Esc to dismiss, Enter to submit
# Floats at horizontal center, ~30% from top of main screen — Spotlight-ish.

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
_panel_session_id: int = 0                  # bumps on each show; late callbacks compare
                                            # to know if their panel-instance is still
                                            # the active one (skip mutation otherwise).
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
_FILE_URL_RE = None  # lazy-compiled below
_ATTACH_EXTS = ("png", "jpg", "jpeg", "gif", "webp", "heic", "pdf")
_MEDIA_TYPES = {
    "png": "image/png",   "jpg": "image/jpeg",
    "jpeg": "image/jpeg", "gif": "image/gif",
    "webp": "image/webp", "heic": "image/heic",
    "pdf": "application/pdf",
}


def _queue_paths_as_images(paths) -> int:
    """Read each path off disk, base64-encode, queue in `_pending_images`,
    update the +N badge on the attach button. Returns count actually added.

    Centralized so the + button, the drag-drop handler, and the URL-strip
    safety net all share one code path. Filters by extension so a stray
    text file dropped on the panel doesn't get encoded.
    """
    global _pending_images
    import base64 as _b64
    from pathlib import Path as _P
    added = 0
    for raw in paths:
        try:
            path = str(raw)
            ext = _P(path).suffix.lower().lstrip(".")
            if ext not in _ATTACH_EXTS:
                continue
            with open(path, "rb") as f:
                data = f.read()
            _pending_images.append({
                "data": _b64.b64encode(data).decode("ascii"),
                "media_type": _MEDIA_TYPES.get(ext, "application/octet-stream"),
                "file_name": _P(path).name,
            })
            added += 1
        except Exception as e:
            sys.stderr.write(f"[tsifl-helper] queue image {raw} failed: {e}\n")
    if added > 0 and _panel_attach_btn is not None:
        try:
            _panel_attach_btn.setTitle_(f"+{len(_pending_images)}")
        except Exception:
            pass
    return added


def _strip_and_extract_paths(text: str):
    """Find file:// URLs in `text` and return (cleaned_text, [file_paths]).

    The user-reported bug: dragging an image onto the input field made
    NSTextView insert the file URL as text (e.g. file:///Users/.../Screenshot.png),
    polluting what the user typed. Custom field editor blocks this at the
    source, but if anything slips through we strip it here as a safety net
    AND auto-queue the file as an attachment — so the user's intent
    (attach image) is preserved.
    """
    import re
    global _FILE_URL_RE
    if _FILE_URL_RE is None:
        # Match file:// URLs (case-insensitive). NSTextView always
        # inserts the full URL with prefix when dropping a file.
        _FILE_URL_RE = re.compile(r"file://\S+", re.IGNORECASE)
    paths: list[str] = []

    def _replace(m):
        url = m.group(0).rstrip(".,;)")  # trim trailing punctuation
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

    cleaned = _FILE_URL_RE.sub(_replace, text)
    # Drop empty lines left behind so the user's typed text doesn't sit
    # under a blank gap where the URL used to be.
    cleaned = "\n".join(line for line in cleaned.splitlines() if line.strip())
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
        (panel_class, target_class, content_view_class, panel_delegate_class)

    All four classes are registered with the Objective-C runtime on first
    call; subsequent panel rebuilds reuse them.
    """
    try:
        from AppKit import NSPanel, NSView, NSTextView
        from Foundation import NSObject
    except Exception as _ie:
        sys.stderr.write(f"[tsifl-helper] AppKit/Foundation unavailable: {_ie}\n")
        return None, None, None, None

    class _TsiflFloatingPanel(NSPanel):
        """Floating panel that CAN become the key window, even when
        borderless. Override required because NSPanel + borderless mask
        defaults to canBecomeKeyWindow=NO, which would block keyboard
        input from reaching the embedded text field."""

        def canBecomeKeyWindow(self):
            return True

        def canBecomeMainWindow(self):
            return False  # menu-bar utility, not the app's main window

    class _TsiflFieldEditor(NSTextView):
        """Custom field editor used by the panel for the input NSTextField.

        Default field editor is a vanilla NSTextView, which accepts file
        URL drops and inserts them as text — that's the user-reported bug
        ('the URL appears in the input when I attach an image'). We
        narrow the acceptable drag types to plain text only AND reject
        the pasteboard read at the lowest level for any file/URL type.

        File drops are still accepted at the panel level (see
        _TsiflContentView below) — they just route to `_pending_images`
        as proper attachments instead of polluting the input text.
        """

        def acceptableDragTypes(self):
            return [
                "public.utf8-plain-text",
                "public.text",
                "NSStringPboardType",
            ]

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

    class _TsiflPanelDelegate(NSObject):
        """NSWindowDelegate that vends `_TsiflFieldEditor` as the field
        editor for the input NSTextField.

        macOS calls `windowWillReturnFieldEditor:toObject:` whenever the
        panel needs to provide a field editor for one of its controls.
        Returning our custom subclass makes the panel use it for the
        input field, which is how we block file URL drops at the source.
        Returning None lets the panel fall back to default behavior for
        any other object (e.g. the response NSTextField, which is
        non-editable so it never asks for a field editor anyway).
        """

        def windowWillReturnFieldEditor_toObject_(self, window, obj):
            try:
                from AppKit import NSTextField
                if not isinstance(obj, NSTextField):
                    return None
                from Foundation import NSMakeRect
                fe = getattr(self, "_field_editor", None)
                if fe is None:
                    fe = _TsiflFieldEditor.alloc().initWithFrame_(
                        NSMakeRect(0, 0, 0, 0)
                    )
                    fe.setFieldEditor_(True)
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
            if self._has_files(sender):
                return 1  # NSDragOperationCopy
            return 0      # NSDragOperationNone

        def draggingUpdated_(self, sender):
            return self.draggingEntered_(sender)

        def prepareForDragOperation_(self, sender):
            return self._has_files(sender)

        def performDragOperation_(self, sender):
            try:
                added = _queue_paths_as_images(self._extract_paths(sender))
                return added > 0
            except Exception as e:
                _log_shortcut_error("drop failed", e)
                return False

        def _has_files(self, sender):
            try:
                return bool(self._extract_paths(sender))
            except Exception:
                return False

        def _extract_paths(self, sender):
            paths: list[str] = []
            try:
                pb = sender.draggingPasteboard()
            except Exception:
                return paths
            # Try the legacy filenames type first — gives plain paths directly
            try:
                files = pb.propertyListForType_("NSFilenamesPboardType")
                if files:
                    return [str(f) for f in files]
            except Exception:
                pass
            # Fall back to URL objects (modern UTI-based reads)
            try:
                from AppKit import NSURL
                urls = pb.readObjectsForClasses_options_([NSURL], None) or []
                for u in urls:
                    p = u.path() if hasattr(u, "path") else None
                    if p:
                        paths.append(str(p))
            except Exception:
                pass
            return paths

    class _PanelTarget(NSObject):
        """Single target object handling Enter (input field action),
        Send button click, and Attach button click. Stateless — reads
        the current `_panel_input` ref from module globals each time."""

        def inputDidEnter_(self, sender):
            if _panel_busy:
                return  # ignore double-Enter while in flight
            text = str(sender.stringValue() or "").strip()
            if text:
                _panel_submit(text)

        def sendClicked_(self, sender):
            if _panel_busy or _panel_input is None:
                return
            text = str(_panel_input.stringValue() or "").strip()
            if text:
                _panel_submit(text)

        def controlTextDidChange_(self, notification):
            """Safety net for the URL-pollution bug.

            Custom field editor blocks file URL drops at the source, but
            paste / Services / odd UTI translations can still slip a
            file:// URL into the input text. This handler runs on every
            user-initiated change, strips file URLs, and auto-queues
            them as proper attachments so the user's intent is preserved.

            `setStringValue_` does NOT fire this method (only field-
            editor mutations do), so the rewrite below cannot recurse.
            """
            if _panel_input is None or _panel_busy:
                return
            try:
                current = str(_panel_input.stringValue() or "")
                if "file://" not in current.lower():
                    return  # fast path — no URL means nothing to do
                cleaned, paths = _strip_and_extract_paths(current)
                if cleaned != current:
                    _panel_input.setStringValue_(cleaned)
                if paths:
                    _queue_paths_as_images(paths)
            except Exception:
                pass

        def attachClicked_(self, sender):
            """Open NSOpenPanel — let user pick image files. Each
            selected file is base64-encoded and queued in
            `_pending_images` via the shared helper; the next submit
            sends them with the chat request."""
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

    return _TsiflFloatingPanel, _PanelTarget, _TsiflContentView, _TsiflPanelDelegate


(
    _TsiflFloatingPanel,
    _PanelTargetClass,
    _TsiflContentViewClass,
    _TsiflPanelDelegateClass,
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
    # Register for file URL drag types. Both legacy (NSFilenamesPboardType)
    # and modern (public.file-url UTI) — Finder uses the legacy type but
    # other source apps may use the UTI form.
    if _TsiflContentViewClass is not None:
        try:
            content.registerForDraggedTypes_(
                ["NSFilenamesPboardType", "public.file-url"]
            )
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
    content.addSubview_(attach_btn)

    # 4. Input field (between logo and attach button)
    input_x = logo_x + logo_size + 12.0
    input_width = attach_x - input_x - 8.0
    input_height = 24.0
    input_y = (INPUT_ROW_HEIGHT - input_height) / 2.0
    input_field = NSTextField.alloc().initWithFrame_(
        NSMakeRect(input_x, input_y, input_width, input_height)
    )
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
    response_field = NSTextField.alloc().initWithFrame_(NSMakeRect(0, 0, 0, 0))
    response_field.setBezeled_(False)
    response_field.setBordered_(False)
    response_field.setDrawsBackground_(False)
    response_field.setFocusRingType_(NSFocusRingTypeNone)
    response_field.setEditable_(False)
    response_field.setSelectable_(True)
    response_field.setFont_(NSFont.systemFontOfSize_(13.0))
    response_field.setTextColor_(muted)
    response_field.setStringValue_("")
    # Multi-line + word wrap. NSLineBreakByWordWrapping = 0
    NSLineBreakByWordWrapping = 0
    response_field.setUsesSingleLineMode_(False)
    response_field.cell().setWraps_(True)
    response_field.cell().setLineBreakMode_(NSLineBreakByWordWrapping)
    response_field.setHidden_(True)
    content.addSubview_(response_field)

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
_PANEL_WIDTH = 540.0
_RESPONSE_PADDING_X = 18.0   # horizontal padding inside the response area
_RESPONSE_PADDING_TOP = 8.0  # gap between response and input row
_RESPONSE_PADDING_BOTTOM = 12.0
_MIN_RESPONSE_HEIGHT = 36.0   # always show at least one full line + padding
_MAX_RESPONSE_HEIGHT = 360.0  # cap so a 5000-char reply doesn't fill the screen


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

        # Position the response field in the bottom area
        _panel_response.setFrame_(NSMakeRect(
            _RESPONSE_PADDING_X,
            _RESPONSE_PADDING_BOTTOM,
            text_width,
            response_height,
        ))
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
    if not _panel_expanded:
        try:
            _panel_response.setHidden_(True)
            _panel_response.setStringValue_("")
        except Exception:
            pass
        return
    try:
        from Foundation import NSMakeRect
        _panel_response.setHidden_(True)
        _panel_response.setStringValue_("")

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
    global _panel_busy, _panel_thinking_timer
    global _panel_ref, _panel_input, _panel_send_btn, _panel_attach_btn
    global _panel_response, _panel_expanded

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
    _panel_expanded = False


def _panel_is_visible() -> bool:
    """True if the panel is currently on-screen and key/visible."""
    if _panel_ref is None:
        return False
    try:
        return bool(_panel_ref.isVisible())
    except Exception:
        return False


def _panel_submit(text: str):
    """User pressed Enter or clicked Send. Don't dismiss — keep panel open,
    show 'thinking…' inline, fire backend, then show the reply inline.

    If the user dismisses the panel mid-request, the request still runs
    to completion in the background; the result surfaces as a macOS
    notification rather than being silently dropped.
    """
    global _panel_busy, _panel_thinking_timer

    global _pending_images
    if _panel_busy:
        return
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

    # Lock + clear the input field. Placeholder shows what was just sent
    # (truncated) so user sees their query in flight.
    if _panel_input is not None:
        try:
            _panel_input.setStringValue_("")
            _panel_input.setEditable_(False)
            attach_note = f" (+{len(images_to_send)} file)" if images_to_send else ""
            _panel_input.setPlaceholderString_(
                f"→ {text[:80]}{attach_note}"
                + ("…" if len(text) > 80 else "")
            )
        except Exception:
            pass

    # Show first thinking line immediately, rotate every 1.5s
    _panel_show_response(_THINKING_LINES[0])
    line_idx = [1]
    def _rotate(timer):
        if not _panel_busy or _panel_session_id != my_session:
            timer.stop()
            return
        if _panel_response is not None and not _panel_response.isHidden():
            _panel_response.setStringValue_(
                _THINKING_LINES[line_idx[0] % len(_THINKING_LINES)]
            )
        line_idx[0] += 1
    t = rumps.Timer(_rotate, 1.5)
    t.start()
    _panel_thinking_timer = t

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
            global _panel_busy, _panel_thinking_timer
            # Always reset busy + stop rotation regardless of session state
            _panel_busy = False
            if _panel_thinking_timer is not None:
                try: _panel_thinking_timer.stop()
                except Exception: pass
                _panel_thinking_timer = None

            reply_text = (result.get("reply") or "(no reply)").strip()
            actions = result.get("actions") or []
            cu_session = result.get("cu_session_id")

            # Show the FULL reply (with paragraph breaks preserved) — the
            # response field is multi-line wrap and the panel grows to fit.
            # Cap at 1500 chars total so a runaway reply doesn't fill the
            # screen; if longer, ellipsize with a hint.
            display = reply_text
            if len(display) > 1500:
                display = display[:1497] + "…"
            extras = []
            if actions:
                extras.append(f"{len(actions)} action(s) queued")
            if cu_session:
                extras.append("helper running")
            tail = f"\n\n({', '.join(extras)})" if extras else ""

            # Persist the last response so reopening the panel later still
            # shows the answer (user feedback: "when i close, the answer
            # disappears, i dont want that"). Cleared when next query is
            # submitted, never disposed otherwise.
            global _last_response_text
            _last_response_text = display + tail

            # If the panel-instance is still the active one and visible,
            # update inline AND re-enable the input so user can ask a
            # follow-up. Otherwise surface as a notification.
            if _panel_session_id == my_session and _panel_is_visible():
                _panel_show_response(display + tail)
                if _panel_input is not None:
                    try:
                        _panel_input.setEditable_(True)
                        _panel_input.setPlaceholderString_("Ask tsifl anything…")
                        _panel_input.becomeFirstResponder()
                    except Exception:
                        pass
            else:
                # Panel was dismissed. Persisted response will surface on
                # next show. Also fire a notification so the user knows
                # the answer is ready.
                try:
                    rumps.notification(
                        title="tsifl",
                        subtitle=text[:60],
                        message=display + tail,
                    )
                except Exception:
                    pass

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
    if _panel_is_visible():
        _panel_dismiss()
        return

    # Always destroy any lingering state before rebuilding (defensive)
    _panel_dismiss()

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
        # That's the safety-net hook that strips any file:// URL the
        # field editor failed to block, and auto-queues the file as a
        # proper attachment.
        try:
            _panel_input.setDelegate_(_panel_target)
        except Exception as e:
            sys.stderr.write(f"[tsifl-helper] input delegate attach failed: {e}\n")

        # Set the panel's window delegate so it vends our custom field
        # editor (the one that rejects file URL drops at the source).
        # Built lazily once per app lifetime.
        global _panel_delegate
        if _panel_delegate is None and _TsiflPanelDelegateClass is not None:
            try:
                _panel_delegate = _TsiflPanelDelegateClass.alloc().init()
            except Exception as e:
                sys.stderr.write(f"[tsifl-helper] panel delegate init failed: {e}\n")
        if _panel_delegate is not None:
            try:
                _panel_ref.setDelegate_(_panel_delegate)
            except Exception as e:
                sys.stderr.write(f"[tsifl-helper] setDelegate_ failed: {e}\n")

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
