# Installing `tsifl Helper`

The desktop helper handles Excel features that the in-Excel add-in can't reach: **Solver**, **Goal Seek**, **Data Tables**, **Scenario Manager**, **SmartArt**, **PivotTables**, **advanced conditional formatting**, and a few more. It runs invisibly in your menu bar — no Terminal needed.

## Quick install (development build, ad-hoc signed)

These instructions are for the unsigned development build. For end-users post-launch, we'll have a code-signed `.dmg` installer.

### 1. Build the `.app`

```bash
cd desktop-agent
python3 setup.py py2app -A   # alias mode — fast, references your dev env
```

Output: `dist/tsifl Helper.app`

For a **shippable** standalone build (slower, ~200MB, runs without your dev env):

```bash
python3 setup.py py2app   # no -A flag
```

### 2. Install it

```bash
mv "dist/tsifl Helper.app" /Applications/
```

Or just drag it into Applications via Finder.

### 3. First launch

Right-click `tsifl Helper.app` → **Open**. (macOS Gatekeeper blocks unsigned apps on first run — right-click + Open is the override path. Subsequent launches work normally.)

You'll see:

1. A `tsifl` text entry appears in your menu bar (top-right of screen)
2. A first-launch dialog asking if you want auto-start at login

If you click **"Auto-start at login"**, macOS will prompt you for permission to manage Login Items. Approve once → the helper auto-launches forever after.

### 4. Verify it's running

Click the `tsifl` menu bar entry. You should see:

- **Status: connected · up Xm Ys** (green check that the agent is polling the backend)
- **Open Logs** → opens `~/Library/Logs/tsifl-agent.log` in your default app
- **Anthropic Console** → quick link to billing/spend dashboard
- **Quit tsifl Helper**

### 5. Test it works end-to-end

Open Excel with any workbook. In the tsifl panel, send:

> *"On Sheet1, run goal seek to make C5 equal to 1000 by changing C2"*

(With C2 holding a number and C5 holding a formula that depends on C2.)

You should see Excel briefly take focus, the cells update, and the panel reply *"Goal Seek converged: C2 changed from X to Y..."*. No Terminal involvement at any point.

## Using the global shortcut (`⌘⇧T`)

The helper registers a global hotkey: **Cmd + Shift + T** anywhere on your Mac. Press it and a small prompt panel pops up — type a request, hit Send, the helper figures out what app you're focused on (Excel, R, Word, etc.) and routes the request appropriately.

Example: you're reading a spec in Chrome, see a number that needs to go in your model. Hit `⌘⇧T` → type *"in cell B5 of Sheet1, set the value to 1.4 million and format as currency"* → tsifl switches focus to Excel and does it.

If the hotkey doesn't trigger anything: you probably haven't granted Input Monitoring permission yet (see below).

## Permissions

On first run, macOS will ask for three permissions. Grant all:

| Permission | Why | Where to grant |
|---|---|---|
| **Automation** (System Events / Microsoft Excel) | Lets the helper send AppleScript commands to Excel | Auto-prompted on first action |
| **Login Items** management | Lets the helper register itself for auto-start | Auto-prompted on first launch dialog |
| **Input Monitoring** | Lets the helper listen for the global ⌘⇧T hotkey | System Settings → Privacy & Security → Input Monitoring → enable `tsifl Helper` |

If you accidentally deny any: System Settings → Privacy & Security → fix individually. After granting Input Monitoring you must **quit and relaunch tsifl Helper** for the listener to attach.

## Troubleshooting

**"tsifl" doesn't appear in menu bar.** Make sure the app actually launched: `ps aux | grep "tsifl Helper"`. If nothing shows, look at console output: launch from Terminal with `"/Applications/tsifl Helper.app/Contents/MacOS/tsifl Helper"`.

**Agent reports errors in the menu.** Click "Open Logs" — the rotating log at `~/Library/Logs/tsifl-agent.log` has a full trace. Most common issue: missing `.env` file or expired Anthropic API key.

**Want to disable auto-start.** System Settings → General → Login Items → uncheck `tsifl Helper`.

**Want to uninstall completely.**
```bash
# Stop the running app
osascript -e 'tell application "tsifl Helper" to quit'

# Remove the bundle
rm -rf "/Applications/tsifl Helper.app"

# Remove the onboarding marker (optional — only matters for re-install)
rm -rf ~/Library/Application\ Support/tsifl-helper

# Remove logs (optional)
rm ~/Library/Logs/tsifl-agent.log*
```

## Architecture

For curious readers / future contributors:

- `tsifl_helper_app.py` is a [rumps](https://github.com/jaredks/rumps) menu bar app that wraps `agent.py`. The agent runs in a daemon thread; the rumps event loop owns the main thread.
- `setup.py` is the py2app build config. `LSUIElement: True` in the plist makes the app background-only (no dock icon, no Cmd-Tab entry).
- The agent itself is the same `agent.py` that you'd run from Terminal — no fork, no special build path. The .app is purely a packaging convenience.
- Login Item registration uses AppleScript via `osascript`, not a private framework, so it works without bundling extra binaries.
- First-launch onboarding marker lives at `~/Library/Application Support/tsifl-helper/.onboarded` so the dialog only shows once.

## Production checklist (pre-launch, not done yet)

- [ ] Apple Developer account ($99/yr) for code-signing
- [ ] Notarize the .app via `xcrun notarytool submit`
- [ ] Build a `.dmg` installer with a custom background image
- [ ] Replace the menu bar text title with a proper `.icns` icon
- [ ] Add an auto-update mechanism (Sparkle? Or just version-check on launch)
- [ ] Sign + notarize the `.dmg` separately
