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

import os
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

import rumps  # type: ignore[import-untyped]


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
        super().__init__("tsifl", quit_button=None)  # custom quit handler below
        self.title = "tsifl"
        # Fallback icon-less mode for first launch — text-only menu bar entry.
        # Once we ship a proper .icns icon, we'll set self.icon = "icon.icns".
        self.menu = [
            rumps.MenuItem("Status: starting...", callback=None),
            None,  # separator
            rumps.MenuItem("Open Logs", callback=self.on_open_logs),
            rumps.MenuItem("Anthropic Console", callback=self.on_open_console),
            None,
            rumps.MenuItem("Quit tsifl Helper", callback=self.on_quit),
        ]

        # Kick off the agent right after the menu bar app is up
        _start_agent_thread()

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
        if _last_error:
            self.title = "tsifl ●"
            status_text = f"Error: {_last_error[:40]}..."
        elif _agent_thread is None or not _agent_thread.is_alive():
            self.title = "tsifl —"
            status_text = "Status: stopped"
        else:
            self.title = "tsifl"
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
