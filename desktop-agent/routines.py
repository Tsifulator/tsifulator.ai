"""
routines.py — persistent recurring tasks ("routines") for tsifl.

A routine is a saved natural-language prompt + a schedule. The scheduler
fires the prompt at the scheduled times. Examples:

  - "every weekday at 8am, summarize my inbox" → routine with schedule
    "weekdays 08:00" and prompt "summarize my inbox"
  - "every hour, snapshot CAT and DE prices" → schedule "every 60 min"
  - "every Sunday at 9pm, check my upcoming week's calendar" →
    schedule "sundays 21:00"

Storage: ~/.tsifl/routines.json (flat list).
Log:     ~/.tsifl/routines.log (one line per execution).
"""

from __future__ import annotations
import json
import re
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional


_ROUTINES_FILE = Path.home() / ".tsifl" / "routines.json"
_RUN_LOG_FILE = Path.home() / ".tsifl" / "routines.log"


@dataclass
class Routine:
    id: str
    name: str
    prompt: str
    schedule: str             # friendly string like "daily 08:00"
    enabled: bool = True
    created: str = ""
    last_run: Optional[str] = None
    last_status: Optional[str] = None  # "ok" | "error" | None
    last_result: Optional[str] = None  # short summary
    next_run: Optional[str] = None     # ISO timestamp (UTC)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Routine":
        return cls(
            id=d.get("id", str(uuid.uuid4())),
            name=d.get("name", ""),
            prompt=d.get("prompt", ""),
            schedule=d.get("schedule", ""),
            enabled=d.get("enabled", True),
            created=d.get("created", ""),
            last_run=d.get("last_run"),
            last_status=d.get("last_status"),
            last_result=d.get("last_result"),
            next_run=d.get("next_run"),
        )


# ─────────────────────────────────────────────────────────────────────────
# Schedule parsing — friendly strings → next-fire datetime
# ─────────────────────────────────────────────────────────────────────────

_DAY_MAP = {
    "sun": 6, "sunday": 6,
    "mon": 0, "monday": 0,
    "tue": 1, "tuesday": 1, "tues": 1,
    "wed": 2, "wednesday": 2,
    "thu": 3, "thursday": 3, "thurs": 3,
    "fri": 4, "friday": 4,
    "sat": 5, "saturday": 5,
}

_DAY_PLURAL_MAP = {
    "weekdays": [0, 1, 2, 3, 4],
    "weekends": [5, 6],
    "daily": list(range(7)),
    "mondays": [0], "tuesdays": [1], "wednesdays": [2], "thursdays": [3],
    "fridays": [4], "saturdays": [5], "sundays": [6],
}


def _parse_hhmm(s: str) -> Optional[tuple[int, int]]:
    """'08:00' / '8:00' / '8am' / '14:30' → (hour, minute) or None."""
    s = s.strip().lower()
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    if m:
        h, mn = int(m.group(1)), int(m.group(2))
        if 0 <= h <= 23 and 0 <= mn <= 59:
            return h, mn
    m = re.match(r"^(\d{1,2})(am|pm)$", s)
    if m:
        h = int(m.group(1))
        ampm = m.group(2)
        if 1 <= h <= 12:
            if ampm == "am":
                h = 0 if h == 12 else h
            else:
                h = 12 if h == 12 else h + 12
            return h, 0
    m = re.match(r"^(\d{1,2}):(\d{2})\s*(am|pm)$", s)
    if m:
        h, mn = int(m.group(1)), int(m.group(2))
        ampm = m.group(3)
        if 1 <= h <= 12 and 0 <= mn <= 59:
            if ampm == "am":
                h = 0 if h == 12 else h
            else:
                h = 12 if h == 12 else h + 12
            return h, mn
    return None


def parse_schedule(schedule: str) -> Optional[dict]:
    """Parse a friendly schedule string into a structured form.

    Returns dict like {"kind": "interval", "minutes": 60} or
    {"kind": "daily_at", "days": [0,1,2,3,4], "hour": 8, "minute": 0}
    or {"kind": "market_hours_interval", "minutes": 60}.

    Returns None if it can't parse.
    """
    s = (schedule or "").strip().lower()
    if not s:
        return None

    # "hourly" / "every hour"
    if s in ("hourly", "every hour"):
        return {"kind": "interval", "minutes": 60}

    # "every N min(s)" / "every N hour(s)"
    m = re.match(r"^every\s+(\d+)\s*(min(?:ute)?s?|hour?s?|hrs?|h)$", s)
    if m:
        n = int(m.group(1))
        unit = m.group(2)
        if unit.startswith("h"):
            return {"kind": "interval", "minutes": n * 60}
        return {"kind": "interval", "minutes": n}

    # "market hours every N min" — fires only 9:30-16:00 weekdays (NY time)
    m = re.match(r"^market\s+hours?\s+every\s+(\d+)\s*(min(?:ute)?s?|hours?|h)$", s)
    if m:
        n = int(m.group(1))
        unit = m.group(2)
        minutes = n * 60 if unit.startswith("h") else n
        return {"kind": "market_hours_interval", "minutes": minutes}

    # "DAY at HH:MM" e.g. "weekdays at 08:00", "monday at 9am",
    # "daily at 8:30am", "weekday 8am" (accepts singular too)
    pattern = re.compile(
        r"^(?P<day>weekdays?|weekends?|daily|"
        r"mondays?|tuesdays?|wednesdays?|thursdays?|fridays?|saturdays?|sundays?)"
        r"(?:\s+at)?\s+(?P<time>.+?)$"
    )
    m = pattern.match(s)
    if m:
        # Normalize singular → plural form for lookup
        day_key = m.group("day")
        if day_key not in _DAY_PLURAL_MAP and day_key + "s" in _DAY_PLURAL_MAP:
            day_key = day_key + "s"
        days = _DAY_PLURAL_MAP.get(day_key)
        time_part = _parse_hhmm(m.group("time").strip())
        if days is not None and time_part is not None:
            return {"kind": "daily_at", "days": days, "hour": time_part[0], "minute": time_part[1]}

    return None


def compute_next_run(schedule: str, now: Optional[datetime] = None) -> Optional[datetime]:
    """Return the next datetime (UTC) that `schedule` should fire after `now`."""
    if now is None:
        now = datetime.now(timezone.utc)
    spec = parse_schedule(schedule)
    if not spec:
        return None

    kind = spec["kind"]

    if kind == "interval":
        return now + timedelta(minutes=spec["minutes"])

    if kind == "market_hours_interval":
        # Treat "market hours" as 13:30 UTC ≤ t ≤ 20:00 UTC weekdays
        # (roughly 9:30am-4pm Eastern; ignores DST drift but close enough).
        # If next +interval lands in a market window, use it; otherwise
        # schedule for next market open.
        cand = now + timedelta(minutes=spec["minutes"])
        # Roll forward until weekday + market hour
        for _ in range(8):  # max 1 week
            if cand.weekday() < 5:
                if cand.hour < 13 or (cand.hour == 13 and cand.minute < 30):
                    cand = cand.replace(hour=13, minute=30, second=0, microsecond=0)
                elif cand.hour > 20 or cand.hour == 20:
                    # past close — roll to next day's open
                    cand = (cand + timedelta(days=1)).replace(hour=13, minute=30, second=0, microsecond=0)
                else:
                    break
            else:
                # weekend → roll to Monday's open
                days_to_mon = (7 - cand.weekday()) % 7 or 1
                cand = (cand + timedelta(days=days_to_mon)).replace(hour=13, minute=30, second=0, microsecond=0)
        return cand

    if kind == "daily_at":
        days = set(spec["days"])
        hour = spec["hour"]
        minute = spec["minute"]
        # Try today first; if it's already past, roll forward
        candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if candidate <= now or candidate.weekday() not in days:
            candidate = candidate + timedelta(days=1)
        # Walk forward until day matches
        for _ in range(8):
            if candidate.weekday() in days and candidate > now:
                return candidate
            candidate = candidate + timedelta(days=1)
        return None

    return None


# ─────────────────────────────────────────────────────────────────────────
# CRUD
# ─────────────────────────────────────────────────────────────────────────

def load_routines() -> list[Routine]:
    if not _ROUTINES_FILE.exists():
        return []
    try:
        data = json.loads(_ROUTINES_FILE.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return [Routine.from_dict(d) for d in data]
    except (json.JSONDecodeError, OSError):
        pass
    return []


def save_routines(routines: list[Routine]):
    _ROUTINES_FILE.parent.mkdir(parents=True, exist_ok=True)
    _ROUTINES_FILE.write_text(
        json.dumps([r.to_dict() for r in routines], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def create_routine(name: str, prompt: str, schedule: str) -> tuple[bool, str]:
    """Create and save a new routine. Returns (ok, message)."""
    spec = parse_schedule(schedule)
    if not spec:
        return False, (
            f"Couldn't parse schedule '{schedule}'. Try: 'daily 08:00', "
            f"'weekdays 9am', 'every 30 min', 'hourly', 'market hours every 60 min'."
        )

    next_run_dt = compute_next_run(schedule)
    if not next_run_dt:
        return False, f"Couldn't compute next run time for '{schedule}'."

    routines = load_routines()
    r = Routine(
        id=str(uuid.uuid4())[:8],
        name=name or prompt[:40],
        prompt=prompt.strip(),
        schedule=schedule.strip().lower(),
        enabled=True,
        created=datetime.now(timezone.utc).isoformat(timespec="seconds"),
        next_run=next_run_dt.isoformat(timespec="seconds"),
    )
    routines.append(r)
    save_routines(routines)
    return True, (
        f"✅ Created routine `{r.name}` — runs {r.schedule}. "
        f"Next fire: {next_run_dt.strftime('%Y-%m-%d %H:%M UTC')}."
    )


def remove_routine(ident: str) -> tuple[bool, str]:
    """Remove a routine by id or name (case-insensitive substring match)."""
    routines = load_routines()
    ident_low = ident.lower().strip()
    remaining = []
    removed = []
    for r in routines:
        if (r.id.lower() == ident_low
                or r.name.lower() == ident_low
                or ident_low in r.name.lower()):
            removed.append(r)
        else:
            remaining.append(r)
    if not removed:
        return False, f"No routine matched '{ident}'."
    save_routines(remaining)
    return True, f"Removed {len(removed)} routine{'s' if len(removed) > 1 else ''}."


def set_enabled(ident: str, enabled: bool) -> tuple[bool, str]:
    """Toggle enabled state."""
    routines = load_routines()
    ident_low = ident.lower().strip()
    matched = []
    for r in routines:
        if (r.id.lower() == ident_low
                or r.name.lower() == ident_low
                or ident_low in r.name.lower()):
            r.enabled = enabled
            matched.append(r)
    if not matched:
        return False, f"No routine matched '{ident}'."
    save_routines(routines)
    word = "enabled" if enabled else "paused"
    return True, f"{word.capitalize()}: {', '.join(r.name for r in matched)}"


def find_routine(ident: str) -> Optional[Routine]:
    """Find a single routine by id or name."""
    routines = load_routines()
    ident_low = ident.lower().strip()
    for r in routines:
        if (r.id.lower() == ident_low
                or r.name.lower() == ident_low
                or ident_low in r.name.lower()):
            return r
    return None


def list_routines_text() -> str:
    """Pretty-print all routines for display."""
    routines = load_routines()
    if not routines:
        return (
            "No routines set.\n"
            "Examples: 'every weekday at 8am summarize my inbox', "
            "'every hour check market', 'daily 9pm note tomorrow\\'s schedule'."
        )
    lines = ["Routines:"]
    for r in routines:
        status = "▶" if r.enabled else "⏸"
        next_run = r.next_run or "?"
        # Pretty-format the next run if it's a valid ISO
        try:
            dt = datetime.fromisoformat(next_run.replace("Z", "+00:00"))
            next_run = dt.strftime("%a %m-%d %H:%M UTC")
        except Exception:
            pass
        last_bit = ""
        if r.last_run:
            try:
                lr = datetime.fromisoformat(r.last_run.replace("Z", "+00:00"))
                last_bit = f"  last: {lr.strftime('%m-%d %H:%M')} ({r.last_status or '?'})"
            except Exception:
                pass
        lines.append(
            f"  {status} [{r.id}] {r.name} — {r.schedule}  next: {next_run}{last_bit}"
        )
        lines.append(f"        \"{r.prompt[:80]}\"")
    return "\n".join(lines)


def record_run(routine_id: str, status: str, result: str):
    """Update a routine's last_run / last_status / last_result and roll its next_run."""
    routines = load_routines()
    for r in routines:
        if r.id == routine_id:
            r.last_run = datetime.now(timezone.utc).isoformat(timespec="seconds")
            r.last_status = status
            r.last_result = (result or "")[:300]
            nxt = compute_next_run(r.schedule)
            r.next_run = nxt.isoformat(timespec="seconds") if nxt else None
            break
    save_routines(routines)
    # Append to log file
    try:
        _RUN_LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with _RUN_LOG_FILE.open("a", encoding="utf-8") as f:
            stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"{stamp} [{routine_id}] {status} — {(result or '')[:160]}\n")
    except Exception:
        pass


def due_routines(now: Optional[datetime] = None) -> list[Routine]:
    """Return routines whose next_run is in the past (and enabled)."""
    if now is None:
        now = datetime.now(timezone.utc)
    out = []
    for r in load_routines():
        if not r.enabled or not r.next_run:
            continue
        try:
            nxt = datetime.fromisoformat(r.next_run.replace("Z", "+00:00"))
            # Ensure both are timezone-aware
            if nxt.tzinfo is None:
                nxt = nxt.replace(tzinfo=timezone.utc)
        except Exception:
            continue
        if nxt <= now:
            out.append(r)
    return out


# ─────────────────────────────────────────────────────────────────────────
# Intent detection — handle routine commands locally without the backend
# ─────────────────────────────────────────────────────────────────────────

_LIST_PATTERNS = [
    re.compile(r"^(?:list\s+)?routines$", re.IGNORECASE),
    re.compile(r"^my\s+routines$", re.IGNORECASE),
    re.compile(r"^show\s+(?:my\s+)?routines$", re.IGNORECASE),
]

_REMOVE_PATTERNS = [
    re.compile(r"^(?:remove|delete|cancel)\s+routine\s+(.+)$", re.IGNORECASE),
]

_PAUSE_PATTERNS = [
    re.compile(r"^(?:pause|disable|stop)\s+routine\s+(.+)$", re.IGNORECASE),
]

_RESUME_PATTERNS = [
    re.compile(r"^(?:resume|enable|start)\s+routine\s+(.+)$", re.IGNORECASE),
]

# Local creation pattern: "<schedule>, <prompt>" or "<schedule> <prompt>"
# Time tokens are strict (digits ± :MM ± am/pm) so "daily 9am brief me…"
# captures schedule="daily 9am" and prompt="brief me…", not the whole thing.
# Also tolerates a leading "every": "every weekday at 8am ..." works.
_TIME_RE = r"\d{1,2}(?::\d{2})?\s*(?:am|pm)?"
_CREATE_PATTERN = re.compile(
    rf"^(?P<sched>"
    rf"every\s+\d+\s*(?:min(?:ute)?s?|hours?|hrs?|h)|"
    rf"hourly|"
    rf"market\s+hours?\s+every\s+\d+\s*(?:min(?:ute)?s?|hours?|h)|"
    rf"(?:every\s+)?"
    rf"(?:daily|weekdays?|weekends?|"
    rf"mondays?|tuesdays?|wednesdays?|thursdays?|fridays?|saturdays?|sundays?)"
    rf"(?:\s+at)?\s+{_TIME_RE}"
    rf")[,:]?\s+(?P<prompt>.+)$",
    re.IGNORECASE,
)


def _normalize_schedule(s: str) -> str:
    """Strip the conversational "every " prefix from day-based schedules
    so parse_schedule sees just 'weekdays 8am' not 'every weekdays 8am'."""
    s = s.strip().lower()
    # Don't strip "every N min" or "every hour" — those need the "every"
    if re.match(r"^every\s+\d+\s", s) or re.match(r"^every\s+hours?\b", s):
        return s
    if s.startswith("every "):
        return s[6:].strip()
    return s


def check_routine_intent(message: str) -> Optional[str]:
    """If `message` is a routine command, handle it locally and return the
    response string. Otherwise return None (caller should send to backend).
    """
    msg = message.strip()

    for pat in _LIST_PATTERNS:
        if pat.match(msg):
            return list_routines_text()

    for pat in _REMOVE_PATTERNS:
        m = pat.match(msg)
        if m:
            _, msg_text = remove_routine(m.group(1))
            return msg_text

    for pat in _PAUSE_PATTERNS:
        m = pat.match(msg)
        if m:
            _, msg_text = set_enabled(m.group(1), False)
            return msg_text

    for pat in _RESUME_PATTERNS:
        m = pat.match(msg)
        if m:
            _, msg_text = set_enabled(m.group(1), True)
            return msg_text

    # Creation: "every X, do Y" or "daily 8am summarize…"
    m = _CREATE_PATTERN.match(msg)
    if m:
        sched = _normalize_schedule(m.group("sched"))
        prompt = m.group("prompt").strip()
        # Name = first ~40 chars of prompt
        name = prompt[:40] + ("…" if len(prompt) > 40 else "")
        _, msg_text = create_routine(name=name, prompt=prompt, schedule=sched)
        return msg_text

    return None
