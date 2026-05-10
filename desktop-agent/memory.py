"""
memory.py — Local persistent memory + custom shortcuts for tsifl desktop agent.

Two storage systems:
  1. Memory: user facts & preferences that persist across sessions
     - "my work email is nick@company.com"
     - "my boss is called Dave"
     Storage: ~/.tsifl/memory.json

  2. Shortcuts: temporary command aliases the user can set/clear
     - "set dataset as cmd+d"  → pressing cmd+d in tsifl = "open the dataset"
     - "set data as /data"     → typing "/data" in tsifl = run the assigned action
     Storage: ~/.tsifl/shortcuts.json
"""

import json
import re
from pathlib import Path
from datetime import datetime

_MEMORY_FILE = Path.home() / ".tsifl" / "memory.json"
_SHORTCUTS_FILE = Path.home() / ".tsifl" / "shortcuts.json"


# ═══════════════════════════════════════════════════════════════════════════════
# MEMORY — persistent user facts & preferences
# ═══════════════════════════════════════════════════════════════════════════════

def load_memories() -> list[dict]:
    """Load all stored memories. Returns list of {fact, added} dicts."""
    if not _MEMORY_FILE.exists():
        return []
    try:
        data = json.loads(_MEMORY_FILE.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except (json.JSONDecodeError, OSError):
        pass
    return []


def save_memory(fact: str) -> str:
    """Save a new memory fact. Returns confirmation string."""
    _MEMORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    memories = load_memories()

    # Check for duplicates (case-insensitive)
    fact_lower = fact.lower().strip()
    for m in memories:
        if m.get("fact", "").lower().strip() == fact_lower:
            return "I already know that."

    # Update existing facts about the same subject
    # e.g. "my email is X" should replace "my email is Y"
    updated = False
    for i, m in enumerate(memories):
        old_fact = m.get("fact", "")
        # Match "my X is Y" pattern — replace if same subject
        old_match = re.match(r"^(my\s+\S+(?:\s+\S+)?\s+(?:is|are))\s+", old_fact, re.IGNORECASE)
        new_match = re.match(r"^(my\s+\S+(?:\s+\S+)?\s+(?:is|are))\s+", fact.strip(), re.IGNORECASE)
        if old_match and new_match and old_match.group(1).lower() == new_match.group(1).lower():
            memories[i] = {
                "fact": fact.strip(),
                "added": datetime.now().isoformat(timespec="seconds"),
            }
            updated = True
            break

    if not updated:
        memories.append({
            "fact": fact.strip(),
            "added": datetime.now().isoformat(timespec="seconds"),
        })

    _MEMORY_FILE.write_text(
        json.dumps(memories, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return "Updated." if updated else "Got it — I'll remember that."


def forget_memory(keyword: str) -> str:
    """Remove memories matching a keyword. Returns confirmation."""
    memories = load_memories()
    keyword_lower = keyword.lower().strip()
    remaining = [m for m in memories if keyword_lower not in m.get("fact", "").lower()]

    removed = len(memories) - len(remaining)
    if removed == 0:
        return f"No memories matched '{keyword}'."

    _MEMORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    _MEMORY_FILE.write_text(
        json.dumps(remaining, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return f"Forgot {removed} memory{'s' if removed > 1 else ''}."


def list_memories() -> str:
    """Return a formatted string of all memories."""
    memories = load_memories()
    if not memories:
        return "No memories stored yet. Tell me things like 'remember my work email is X'."
    lines = []
    for i, m in enumerate(memories, 1):
        lines.append(f"  {i}. {m['fact']}")
    return f"I remember {len(memories)} thing{'s' if len(memories) > 1 else ''}:\n" + "\n".join(lines)


def get_memory_context() -> str:
    """Get memories formatted for injection into the system prompt context."""
    memories = load_memories()
    shortcuts = load_shortcuts()
    parts = []

    if memories:
        facts = [m["fact"] for m in memories]
        parts.append("User preferences & facts (remember these):\n" + "\n".join(f"- {f}" for f in facts))

    if shortcuts:
        sc_lines = [f"- /{s['trigger']} → \"{s['action']}\"" for s in shortcuts]
        parts.append("User shortcuts (execute the action when user types the trigger):\n" + "\n".join(sc_lines))

    return "\n\n".join(parts)


# ═══════════════════════════════════════════════════════════════════════════════
# SHORTCUTS — temporary command aliases
# ═══════════════════════════════════════════════════════════════════════════════

def load_shortcuts() -> list[dict]:
    """Load all stored shortcuts. Returns list of {trigger, action, description, added} dicts."""
    if not _SHORTCUTS_FILE.exists():
        return []
    try:
        data = json.loads(_SHORTCUTS_FILE.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except (json.JSONDecodeError, OSError):
        pass
    return []


def save_shortcut(trigger: str, action: str, description: str = "",
                  hotkey: str = "") -> str:
    """Save or update a custom shortcut.

    Args:
        trigger: the keyword/slash command (e.g. "data", "report", "emails")
        action: what to do when triggered (e.g. "open ~/Desktop/data.csv in RStudio")
        description: optional human-readable description
        hotkey: optional system hotkey combo (e.g. "cmd+d") — registers a global shortcut
    """
    _SHORTCUTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    shortcuts = load_shortcuts()
    trigger_clean = trigger.lower().strip().lstrip("/")

    entry = {
        "trigger": trigger_clean,
        "action": action.strip(),
        "description": description or action.strip()[:60],
        "added": datetime.now().isoformat(timespec="seconds"),
    }
    if hotkey:
        entry["hotkey"] = hotkey.strip().lower()

    # Update if trigger already exists
    for i, s in enumerate(shortcuts):
        if s.get("trigger", "").lower() == trigger_clean:
            shortcuts[i] = entry
            _SHORTCUTS_FILE.write_text(
                json.dumps(shortcuts, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            # Register the global hotkey if specified
            hk_msg = ""
            if hotkey:
                hk_msg = _try_register_hotkey(trigger_clean, action.strip(), hotkey)
            return f"Updated /{trigger_clean}" + (f" ({hk_msg})" if hk_msg else "")

    shortcuts.append(entry)
    _SHORTCUTS_FILE.write_text(
        json.dumps(shortcuts, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    # Register the global hotkey if specified
    hk_msg = ""
    if hotkey:
        hk_msg = _try_register_hotkey(trigger_clean, action.strip(), hotkey)

    base = f"Created /{trigger_clean}"
    if hk_msg:
        return f"{base} ({hk_msg})"
    return f"{base} — type it anytime to run: {action.strip()[:60]}"


def _try_register_hotkey(trigger: str, action: str, hotkey: str) -> str:
    """Try to register a system-level hotkey. Returns status message."""
    try:
        from tsifl_helper_app import _register_dynamic_hotkey
        ok, msg = _register_dynamic_hotkey(trigger, action, hotkey)
        return msg
    except Exception as e:
        return f"hotkey registration failed: {e}"


def remove_shortcut(trigger: str) -> str:
    """Remove a shortcut by trigger name."""
    shortcuts = load_shortcuts()
    trigger_clean = trigger.lower().strip().lstrip("/")
    remaining = [s for s in shortcuts if s.get("trigger", "").lower() != trigger_clean]

    if len(remaining) == len(shortcuts):
        return f"No shortcut '/{trigger_clean}' found."

    _SHORTCUTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    _SHORTCUTS_FILE.write_text(
        json.dumps(remaining, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return f"Removed /{trigger_clean}"


def list_shortcuts() -> str:
    """Return a formatted string of all shortcuts."""
    shortcuts = load_shortcuts()
    if not shortcuts:
        return "No shortcuts set.\nSay 'set X as /name' or 'set X as cmd+D' to create one."
    lines = []
    for s in shortcuts:
        hotkey = s.get("hotkey", "")
        if hotkey:
            pretty = _pretty_hotkey(hotkey)
            lines.append(f"  {pretty} (/{s['trigger']}) → {s['action'][:60]}")
        else:
            lines.append(f"  /{s['trigger']} → {s['action'][:80]}")
    return f"Your shortcuts:\n" + "\n".join(lines)


def _pretty_hotkey(combo: str) -> str:
    """Turn 'cmd+d' into '⌘D' for display."""
    symbols = {"cmd": "⌘", "command": "⌘", "shift": "⇧", "opt": "⌥",
               "option": "⌥", "alt": "⌥", "ctrl": "⌃", "control": "⌃"}
    parts = [p.strip().lower() for p in combo.replace("+", " ").split()]
    result = ""
    key = ""
    for p in parts:
        if p in symbols:
            result += symbols[p]
        else:
            key = p.upper()
    return result + key


def resolve_shortcut(message: str) -> str | None:
    """Check if the message matches a shortcut trigger. Returns the action or None.

    Matches: /trigger, /trigger with extra text
    The action text is returned so it can be sent to the backend as if the user typed it.
    """
    msg = message.strip()
    if not msg.startswith("/"):
        return None

    # Extract the trigger word
    parts = msg[1:].split(None, 1)
    if not parts:
        return None
    trigger = parts[0].lower()

    shortcuts = load_shortcuts()
    for s in shortcuts:
        if s.get("trigger", "").lower() == trigger:
            action = s["action"]
            # If user typed extra text after the trigger, append it
            if len(parts) > 1:
                action += " " + parts[1]
            return action

    return None  # no matching shortcut


def clear_all_shortcuts() -> str:
    """Remove all shortcuts."""
    if _SHORTCUTS_FILE.exists():
        _SHORTCUTS_FILE.unlink()
    return "All shortcuts cleared."


# ═══════════════════════════════════════════════════════════════════════════════
# INTENT DETECTION — routes memory/shortcut commands locally
# ═══════════════════════════════════════════════════════════════════════════════

_REMEMBER_PATTERNS = [
    re.compile(r"^remember\s+(?:that\s+)?(.+)", re.IGNORECASE),
    re.compile(r"^note\s+(?:that\s+)?(.+)", re.IGNORECASE),
    re.compile(r"^save\s+(?:that\s+)?(.+)", re.IGNORECASE),
    re.compile(r"^my\s+(.+?\s+(?:is|are)\s+.+)", re.IGNORECASE),
]

_FORGET_PATTERNS = [
    re.compile(r"^forget\s+(?:about\s+)?(.+)", re.IGNORECASE),
    re.compile(r"^delete\s+memory\s+(?:about\s+)?(.+)", re.IGNORECASE),
    re.compile(r"^remove\s+memory\s+(?:about\s+)?(.+)", re.IGNORECASE),
]

_LIST_PATTERNS = [
    re.compile(r"^(?:what do you|what you)\s+(?:remember|know)\s*(?:about me)?", re.IGNORECASE),
    re.compile(r"^(?:list|show)\s+(?:my\s+)?memories", re.IGNORECASE),
    re.compile(r"^memories$", re.IGNORECASE),
]

# Shortcut patterns: "set X as /name", "set X as cmd+d", etc.
_SET_SHORTCUT_PATTERNS = [
    # "set <action> as /trigger"
    re.compile(r"^set\s+(.+?)\s+as\s+/(\S+)$", re.IGNORECASE),
    # "create shortcut /trigger for <action>"
    re.compile(r"^(?:create|add)\s+(?:shortcut|command)\s+/(\S+)\s+(?:for|to|=)\s+(.+)$", re.IGNORECASE),
    # "shortcut /trigger = action"
    re.compile(r"^(?:shortcut|cmd)\s+/(\S+)\s*=\s*(.+)$", re.IGNORECASE),
]

# System hotkey patterns: "set <action> as cmd+d", "bind cmd+d to <action>"
_SET_HOTKEY_PATTERNS = [
    # "set <action> as cmd+d"
    re.compile(r"^set\s+(.+?)\s+as\s+((?:cmd|command|ctrl|control|opt|option|alt|shift)[\+\s]\S+)$", re.IGNORECASE),
    # "bind cmd+d to <action>"
    re.compile(r"^bind\s+((?:cmd|command|ctrl|control|opt|option|alt|shift)[\+\s]\S+)\s+(?:to|=)\s+(.+)$", re.IGNORECASE),
    # "hotkey cmd+d = action"
    re.compile(r"^hotkey\s+((?:cmd|command|ctrl|control|opt|option|alt|shift)[\+\s]\S+)\s*=\s*(.+)$", re.IGNORECASE),
]

_REMOVE_SHORTCUT_PATTERNS = [
    re.compile(r"^(?:remove|delete|clear)\s+(?:shortcut|command)\s+/(\S+)$", re.IGNORECASE),
    re.compile(r"^unset\s+/(\S+)$", re.IGNORECASE),
]

_LIST_SHORTCUTS_PATTERNS = [
    re.compile(r"^(?:list|show)\s+(?:my\s+)?(?:shortcuts|commands)$", re.IGNORECASE),
    re.compile(r"^shortcuts$", re.IGNORECASE),
    re.compile(r"^my\s+shortcuts$", re.IGNORECASE),
]

_CLEAR_SHORTCUTS_PATTERNS = [
    re.compile(r"^clear\s+(?:all\s+)?shortcuts$", re.IGNORECASE),
]


def check_memory_intent(message: str) -> str | None:
    """Check if the message is a memory or shortcut command. Returns response string or None.

    Handles locally — these never need to go to the backend.
    """
    msg = message.strip()

    # ── Shortcut commands ─────────────────────────────────
    # List shortcuts
    for pat in _LIST_SHORTCUTS_PATTERNS:
        if pat.match(msg):
            return list_shortcuts()

    # Clear all shortcuts
    for pat in _CLEAR_SHORTCUTS_PATTERNS:
        if pat.match(msg):
            return clear_all_shortcuts()

    # Remove shortcut
    for pat in _REMOVE_SHORTCUT_PATTERNS:
        m = pat.match(msg)
        if m:
            return remove_shortcut(m.group(1))

    # Set system hotkey — "set X as cmd+d"
    for pat in _SET_HOTKEY_PATTERNS:
        m = pat.match(msg)
        if m:
            groups = m.groups()
            if len(groups) == 2:
                if pat == _SET_HOTKEY_PATTERNS[0]:
                    action, combo = groups
                else:
                    combo, action = groups
                # Use the key letter as the trigger name
                key_part = combo.replace("+", " ").split()[-1].lower()
                trigger = key_part if len(key_part) <= 3 else key_part[:3]
                return save_shortcut(trigger, action, hotkey=combo)

    # Set slash shortcut — "set X as /name"
    for pat in _SET_SHORTCUT_PATTERNS:
        m = pat.match(msg)
        if m:
            groups = m.groups()
            if len(groups) == 2:
                # Figure out which group is trigger vs action
                # Pattern 1: (action, trigger) — "set <action> as /<trigger>"
                # Pattern 2&3: (trigger, action) — "create shortcut /<trigger> for <action>"
                if pat == _SET_SHORTCUT_PATTERNS[0]:
                    action, trigger = groups
                else:
                    trigger, action = groups
                return save_shortcut(trigger, action)

    # ── Memory commands ───────────────────────────────────
    # List memories
    for pat in _LIST_PATTERNS:
        if pat.match(msg):
            return list_memories()

    # Forget
    for pat in _FORGET_PATTERNS:
        m = pat.match(msg)
        if m:
            return forget_memory(m.group(1))

    # Remember
    for pat in _REMEMBER_PATTERNS:
        m = pat.match(msg)
        if m:
            return save_memory(m.group(1))

    return None  # not a memory/shortcut command — proceed normally
