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

# Global hotkey listening: we use AppKit's NSEvent.addGlobalMonitorForEvents
# directly via PyObjC. This is the same path Spotlight, Raycast, Alfred etc.
# all use on macOS — it's the supported, stable API and works regardless of
# whether the .app is code-signed or rebuilt with a different bundle hash.
#
# We previously tried `pynput` here. pynput uses CGEventTap under the hood
# and depends on the Input Monitoring permission propagating cleanly to the
# bundled .app — a flaky path on macOS Sonoma+, especially with unsigned
# builds. NSEvent's global monitor is a higher-level API on the same OS
# event system but with much better permission semantics.
try:
    from AppKit import NSEvent  # type: ignore[import-untyped]
    _NSEVENT_AVAILABLE = True
except Exception:
    NSEvent = None  # type: ignore[assignment]
    _NSEVENT_AVAILABLE = False

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


def _send_to_backend(message: str, frontmost_app: str) -> dict:
    """POST the user's prompt to the backend /chat endpoint. Synchronous —
    called from the main UI thread after the panel collects input. Returns
    the parsed JSON response or an error-shaped dict."""
    if not _HTTPX_AVAILABLE:
        return {"reply": "Helper is missing httpx; reinstall the bundle.",
                "actions": []}

    backend = os.getenv(
        "BACKEND_URL",
        "https://focused-solace-production-6839.up.railway.app",
    )
    # Map macOS frontmost app names to backend's `app` field. Default to
    # excel since that's the most common case; the addin handles routing.
    app_map = {
        "Microsoft Excel": "excel",
        "RStudio": "rstudio",
        "Microsoft PowerPoint": "powerpoint",
        "Microsoft Word": "word",
    }
    app = app_map.get(frontmost_app, "shortcut")

    body = {
        "user_id": "shortcut-anon",  # TODO: thread real user_id once auth lands
        "message": message,
        "context": {
            "app": app,
            "source": "global-shortcut",
            "frontmost_app": frontmost_app,
        },
    }
    try:
        with _httpx.Client(timeout=60) as client:
            r = client.post(f"{backend.rstrip('/')}/chat/", json=body)
            if r.status_code == 200:
                return r.json()
            return {"reply": f"Backend error ({r.status_code}): {r.text[:200]}",
                    "actions": []}
    except Exception as e:
        return {"reply": f"Could not reach tsifl: {e}", "actions": []}


def _show_shortcut_panel():
    """Pop a modal input panel, capture the user's prompt, send to backend,
    show the reply. Called on the rumps main thread (dispatched from the
    pynput hotkey callback).

    Uses rumps.Window — a native macOS NSAlert-style dialog. Modal, but
    that's fine for a one-shot shortcut. Future polish: replace with a
    non-modal NSPanel for that Spotlight-style float.
    """
    frontmost = _detect_frontmost_app()
    prompt_label = f"Working in {frontmost}" if frontmost not in ("unknown",) else "What can tsifl do?"

    win = rumps.Window(
        title="tsifl",
        message=prompt_label,
        ok="Send",
        cancel="Cancel",
        dimensions=(440, 80),
    )
    response = win.run()
    if not response.clicked or not response.text.strip():
        return

    # Show a brief "working" indicator, then make the call. Doing this on
    # the UI thread keeps it simple — the user already paused for the
    # modal so a few-second sync request is acceptable.
    try:
        result = _send_to_backend(response.text.strip(), frontmost)
    except Exception as e:
        result = {"reply": f"Request failed: {e}", "actions": []}

    reply_text = result.get("reply") or "(no reply)"
    actions = result.get("actions") or []
    cu_session = result.get("cu_session_id")

    # Build the response window message. If actions were emitted, mention
    # they're queued; the addin or agent will execute them.
    parts = [reply_text]
    if actions:
        types = ", ".join(sorted({a.get("type", "?") for a in actions}))
        parts.append(f"\n\n→ {len(actions)} action(s) queued: {types}")
    if cu_session:
        parts.append(f"\n\n→ Helper session: {cu_session}")
    final = "\n".join(parts)

    rumps.alert(
        title="tsifl",
        message=final[:1500],  # macOS alerts truncate hard, keep it sane
        ok="OK",
    )


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


# NSEvent monitor handle. Kept module-level so it doesn't get GC'd —
# PyObjC monitor refs are weak; if the Python ref is dropped, the monitor
# stops firing.
_event_monitor: "object | None" = None
_event_monitor_local: "object | None" = None

# AppKit constants. NSEvent flag values are stable across macOS versions
# (these are documented public APIs).
_NS_EVENT_MASK_KEY_DOWN = 1 << 10                  # NSEventMaskKeyDown
_NS_EVENT_MODIFIER_FLAG_COMMAND = 1 << 20          # NSEventModifierFlagCommand
_NS_EVENT_MODIFIER_FLAG_SHIFT = 1 << 17            # NSEventModifierFlagShift


def _ns_event_handler(event):
    """Called by AppKit when a key-down event fires globally (any app).

    Returns nothing to AppKit (global monitor doesn't consume events; the
    keystroke still passes through to whatever app was focused). For our
    purposes that's fine — ⌘⇧T isn't a system shortcut, so nothing else
    will react to it.
    """
    try:
        flags = event.modifierFlags()
        # Filter: only act when BOTH Cmd and Shift are pressed.
        if not (flags & _NS_EVENT_MODIFIER_FLAG_COMMAND):
            return
        if not (flags & _NS_EVENT_MODIFIER_FLAG_SHIFT):
            return
        chars = event.charactersIgnoringModifiers()
        if chars and str(chars).lower() == "t":
            _on_shortcut_pressed()
    except Exception as e:
        sys.stderr.write(f"[tsifl-helper] NSEvent handler error: {e}\n")


def _start_hotkey_listener() -> bool:
    """Register a global Cmd+Shift+T listener via AppKit's NSEvent API.

    Returns True if registration succeeded. Note that "succeeded" here just
    means AppKit returned a non-nil monitor — if Input Monitoring permission
    isn't granted, the monitor object exists but never fires. That's the
    same trap pynput hit; the only mitigation is to ensure permission is
    granted (which the user has now done) and to relaunch the app fresh
    so the system picks up the entitlement.

    We also register a LOCAL monitor (only fires when our app is frontmost,
    which we never are since we're a menu bar app) — this is mostly
    defensive: if the global monitor's permission scope ever changes, the
    local one might still work for in-app testing.
    """
    global _event_monitor, _event_monitor_local
    if not _NSEVENT_AVAILABLE:
        sys.stderr.write("[tsifl-helper] AppKit/NSEvent not available\n")
        return False
    try:
        _event_monitor = NSEvent.addGlobalMonitorForEventsMatchingMask_handler_(
            _NS_EVENT_MASK_KEY_DOWN, _ns_event_handler,
        )
        _event_monitor_local = NSEvent.addLocalMonitorForEventsMatchingMask_handler_(
            _NS_EVENT_MASK_KEY_DOWN,
            lambda event: (_ns_event_handler(event), event)[1],  # local must return event
        )
        if _event_monitor is None:
            sys.stderr.write(
                "[tsifl-helper] NSEvent.addGlobalMonitor returned nil. "
                "Likely Input Monitoring permission still not propagated to this build.\n"
            )
            return False
        return True
    except Exception as e:
        sys.stderr.write(f"[tsifl-helper] NSEvent monitor failed to attach: {e}\n")
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
            rumps.MenuItem("Open with ⌘⇧T", callback=self.on_open_shortcut),
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
                "[tsifl-helper] global ⌘⇧T listener could not attach. Verify "
                "Input Monitoring is ON for tsifl Helper in System Settings → "
                "Privacy & Security, then quit and relaunch the app.\n"
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
