"""
memory.py — Local persistent memory for tsifl desktop agent.

Stores user preferences, facts, and context that persist across sessions.
Examples:
  - "my work email is nick@company.com"
  - "use Safari for personal browsing"
  - "my boss is called Dave"

Storage: ~/.tsifl/memory.json — simple flat list of facts.
The full list is sent as context with every request so Claude can use it.
"""

import json
import re
from pathlib import Path
from datetime import datetime

_MEMORY_FILE = Path.home() / ".tsifl" / "memory.json"


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

    memories.append({
        "fact": fact.strip(),
        "added": datetime.now().isoformat(timespec="seconds"),
    })
    _MEMORY_FILE.write_text(
        json.dumps(memories, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return f"Got it — I'll remember that."


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
    if not memories:
        return ""
    facts = [m["fact"] for m in memories]
    return "User preferences & facts (remember these):\n" + "\n".join(f"- {f}" for f in facts)


# ── Intent detection: does the user want to save/forget/list memories? ────

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


def check_memory_intent(message: str) -> str | None:
    """Check if the message is a memory command. Returns response string or None.

    Handles locally — these never need to go to the backend.
    """
    msg = message.strip()

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

    return None  # not a memory command — proceed normally
