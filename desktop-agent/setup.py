"""
py2app build config for tsifl Helper.app.

Build:
    cd desktop-agent
    python3 setup.py py2app

Output:
    dist/tsifl Helper.app  ← drag this to /Applications

The app is a background-only menu bar app (LSUIElement=True, no dock
icon, no main window). It runs the agent's poll loop in a thread and
surfaces status via a menu bar entry.

Notes for first ship:
  - Unsigned. Gatekeeper will warn on first launch; user must
    right-click → Open the first time, OR System Settings → Security &
    Privacy → "Open Anyway." Code-signing requires an Apple Developer
    account ($99/yr) — defer until pre-launch.
  - LoginItem registration is NOT done at build time — that requires
    user permission on first launch (handled in tsifl_helper_app.py
    via a one-time prompt, post-MVP).
"""

from setuptools import setup

# The script that becomes the app's main entry point
APP = ["tsifl_helper_app.py"]

# Files that need to be included in the bundle's Resources directory.
# Includes the agent's sibling modules — they're imported at runtime.
DATA_FILES = [
    "agent.py",
    "excel_applescript.py",
    "services_local.py",
]

OPTIONS = {
    "argv_emulation": False,
    "plist": {
        # ── Bundle identity ─────────────────────────────────────────────
        "CFBundleName":            "tsifl Helper",
        "CFBundleDisplayName":     "tsifl Helper",
        "CFBundleIdentifier":      "ai.tsifulator.helper",
        "CFBundleVersion":         "0.1.0",
        "CFBundleShortVersionString": "0.1.0",

        # ── Background-only app (no dock icon, no main window) ─────────
        # Per Apple docs: LSUIElement=True turns this into a "menu bar
        # accessory" — the app runs but doesn't show in Cmd+Tab or Dock.
        # Exactly the UX we want for a helper agent.
        "LSUIElement":             True,

        # ── Permissions usage descriptions (required for macOS perms) ──
        # The agent uses AppleScript and (for fallback CU) screen capture.
        # These strings appear in the System Settings prompts when macOS
        # asks the user to grant Accessibility / Automation / Screen
        # Recording permissions.
        "NSAppleEventsUsageDescription": (
            "tsifl Helper needs to control Microsoft Excel via AppleScript "
            "to run advanced features like Goal Seek, Solver, Data Tables, "
            "and SmartArt that aren't available through the Excel add-in API."
        ),
        "NSAccessibilityUsageDescription": (
            "tsifl Helper uses Accessibility for the rare cases where it "
            "needs to interact with Excel's UI directly. Most operations "
            "use AppleScript and don't require this."
        ),
        # Minimum macOS version. py2app handles the actual SDK check.
        "LSMinimumSystemVersion":  "12.0",
    },
    # Python packages we explicitly include. Some of these have C extensions
    # or runtime imports that py2app's default scanner misses.
    "packages": [
        "rumps",
        "anthropic",
        "httpx",
        "dotenv",
        "xlwings",
        "PIL",
        "appscript",     # xlwings's Mac backend
    ],
    # Modules we explicitly include even if py2app doesn't auto-detect them.
    "includes": [
        "agent",
        "excel_applescript",
        "services_local",
    ],
    # Files / dirs to exclude — these would bloat the bundle without being needed.
    "excludes": [
        "tkinter",
        "test",
        "tests",
        "unittest",
        "pydoc",
        "doctest",
    ],
    # Bundling strategy:
    #   - "alias" mode = .app references the dev environment (fast iterate, won't ship)
    #   - default     = .app contains a self-contained Python (slow build, shippable)
    # We default to shippable. Pass `--alias` on the build command for fast dev.
}

setup(
    app=APP,
    name="tsifl Helper",
    data_files=DATA_FILES,
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
)
