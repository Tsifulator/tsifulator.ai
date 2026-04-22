"""RStudio rubric evaluator.

R actions come back shaped very differently from Excel — typically ONE
`run_r_code` action with the entire script inside `payload.code`. So the
checks here are mostly about what's inside that code string, not about
cell-level writes.

Per the SYSTEM_PROMPT rule: R responses must be exactly ONE run_r_code
action with all code combined. Multiple run_r_code actions in one turn
are a regression we want to catch.
"""
from __future__ import annotations

import re
from typing import Callable


# ── Helpers ──────────────────────────────────────────────────────────────────

def _combined_code(actions: list[dict]) -> str:
    """Concatenate all `payload.code` strings from run_r_code actions."""
    chunks = []
    for a in actions:
        if a.get("type") == "run_r_code":
            code = (a.get("payload") or {}).get("code") or ""
            if isinstance(code, str):
                chunks.append(code)
    return "\n".join(chunks)


# ── Rubric check functions ───────────────────────────────────────────────────

def must_include_action_types(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    present = {a.get("type") for a in actions}
    missing = [t for t in spec if t not in present]
    if missing:
        return False, f"missing action types: {missing} (present: {sorted(present)})"
    return True, f"all required types present: {spec}"


def max_action_count(actions: list[dict], spec: int) -> tuple[bool, str]:
    if len(actions) > spec:
        return False, f"got {len(actions)} actions, max allowed {spec}"
    return True, f"{len(actions)} actions (≤ {spec})"


def min_action_count(actions: list[dict], spec: int) -> tuple[bool, str]:
    if len(actions) < spec:
        return False, f"got {len(actions)} actions, need ≥ {spec}"
    return True, f"{len(actions)} actions (≥ {spec})"


def exactly_one_run_r_code(actions: list[dict], spec: bool) -> tuple[bool, str]:
    """SYSTEM_PROMPT rule: R responses must be ONE combined run_r_code.
    Multiple = regression. Zero = also a fail.
    """
    count = sum(1 for a in actions if a.get("type") == "run_r_code")
    if spec and count != 1:
        return False, f"expected exactly 1 run_r_code action, got {count}"
    return True, f"run_r_code count = {count}"


def code_contains(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    code = _combined_code(actions)
    missing = [s for s in spec if s not in code]
    if missing:
        return False, f"code missing substrings: {missing}"
    return True, "all required code substrings present"


def code_contains_regex(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    code = _combined_code(actions)
    failed = [p for p in spec if not re.search(p, code)]
    if failed:
        return False, f"code missing regex patterns: {failed}"
    return True, "all required regex patterns matched"


def code_forbids(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    code = _combined_code(actions)
    hits = [s for s in spec if s in code]
    if hits:
        return False, f"code contains forbidden substrings: {hits}"
    return True, "no forbidden substrings in code"


def reply_not_empty(actions: list[dict], spec: bool, reply: str = "",
                    cu_session_id: str | None = None) -> tuple[bool, str]:
    has_reply = bool((reply or "").strip())
    has_cu = bool(cu_session_id)
    ok = has_reply or has_cu
    if spec and not ok:
        return False, "reply is empty and no cu_session_id"
    return True, "reply present" if has_reply else ("cu session active" if has_cu else "no reply required")


# ── Dispatcher ───────────────────────────────────────────────────────────────

CHECKS: dict[str, Callable[..., tuple[bool, str]]] = {
    "must_include_action_types": must_include_action_types,
    "max_action_count":           max_action_count,
    "min_action_count":           min_action_count,
    "exactly_one_run_r_code":     exactly_one_run_r_code,
    "code_contains":              code_contains,
    "code_contains_regex":        code_contains_regex,
    "code_forbids":               code_forbids,
    "reply_not_empty":            reply_not_empty,
}


def evaluate(rubric: dict, actions: list[dict], reply: str = "",
             cu_session_id: str | None = None) -> list[tuple[bool, str, str]]:
    results: list[tuple[bool, str, str]] = []
    for name, spec in rubric.items():
        if name.startswith("_"):
            continue
        check = CHECKS.get(name)
        if check is None:
            results.append((False, name, "UNKNOWN RUBRIC KEY"))
            continue
        try:
            if name == "reply_not_empty":
                ok, detail = check(actions, spec, reply=reply, cu_session_id=cu_session_id)
            else:
                ok, detail = check(actions, spec)
        except Exception as e:
            ok, detail = False, f"check raised: {type(e).__name__}: {e}"
        results.append((ok, name, detail))
    return results
