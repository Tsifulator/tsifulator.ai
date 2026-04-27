"""Excel rubric evaluator.

Takes the `actions` list returned by /chat plus the `rubric.yaml` for a case,
returns a list of (pass, name, detail) tuples. Every rubric key maps to a
function here; unknown keys raise so typos are caught loudly.
"""
from __future__ import annotations

import re
from typing import Any, Callable


# ── Helpers ──────────────────────────────────────────────────────────────────

def _iter_action_sheets(actions: list[dict]) -> list[str]:
    """Every sheet name referenced anywhere in the action batch."""
    out: list[str] = []
    for a in actions:
        p = a.get("payload") or {}
        sheet = p.get("sheet")
        if isinstance(sheet, str) and sheet:
            out.append(sheet)
        ref = p.get("reference")
        if isinstance(ref, str) and "!" in ref:
            out.append(ref.split("!", 1)[0].strip().strip("'"))
    return out


def _cell_addr(sheet: str | None, cell: str) -> str:
    return f"{sheet}!{cell}" if sheet else cell


def _action_targets(actions: list[dict]) -> list[str]:
    """Normalized list of `Sheet!Cell` addresses written or formatted."""
    out: list[str] = []
    for a in actions:
        t = a.get("type", "")
        p = a.get("payload") or {}
        s = p.get("sheet")
        if t in ("write_cell", "write_formula"):
            cell = p.get("cell") or p.get("address")
            if cell:
                out.append(_cell_addr(s, cell))
        elif t in ("write_range", "clear_range", "format_range", "fill_down", "fill_right"):
            rng = p.get("range")
            if rng:
                out.append(_cell_addr(s, rng))
        elif t == "create_named_range":
            ref = p.get("reference") or ""
            if ref:
                out.append(ref)
    return out


def _cell_in_range(cell: str, rng: str) -> bool:
    """True if `cell` (e.g. "A1" or "Sheet!A1") is inside `rng` (e.g. "Sheet!A1:C10")."""
    def split(x: str) -> tuple[str | None, str]:
        if "!" in x:
            s, c = x.split("!", 1)
            return s.strip().strip("'"), c
        return None, x

    c_sheet, c_cell = split(cell)
    r_sheet, r_body = split(rng)
    if r_sheet and c_sheet and r_sheet.casefold() != c_sheet.casefold():
        return False
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", r_body)
    m2 = re.match(r"^\$?([A-Z]+)\$?(\d+)$", c_cell)
    if not m or not m2:
        return False
    c_col, c_row = m2.group(1), int(m2.group(2))
    r_c1, r_r1 = m.group(1), int(m.group(2))
    r_c2 = m.group(3) or r_c1
    r_r2 = int(m.group(4)) if m.group(4) else r_r1
    return _col_num(r_c1) <= _col_num(c_col) <= _col_num(r_c2) and r_r1 <= c_row <= r_r2


def _col_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _find_action_at(actions: list[dict], sheet: str, cell: str) -> dict | None:
    for a in actions:
        p = a.get("payload") or {}
        s = p.get("sheet")
        c = p.get("cell") or p.get("address")
        if s == sheet and c == cell:
            return a
    return None


# ── Rubric check functions ───────────────────────────────────────────────────

def must_include_action_types(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    present = {a.get("type") for a in actions}
    missing = [t for t in spec if t not in present]
    if missing:
        return False, f"missing action types: {missing} (present: {sorted(present)})"
    return True, f"all required types present: {spec}"


def must_reference_cells(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    targets = _action_targets(actions)
    missing = []
    for required in spec:
        # required can be "Sheet!Cell" or "Sheet!Range"; at least one target must be inside it
        if not any(_cell_in_range(t, required) or _cell_in_range(required, t) or t == required for t in targets):
            missing.append(required)
    if missing:
        return False, f"no actions targeting: {missing}"
    return True, f"all required cells referenced"


def formula_at_cell_contains(actions: list[dict], spec: dict[str, list[str]]) -> tuple[bool, str]:
    failures = []
    for addr, required_substrings in spec.items():
        sheet, cell = addr.split("!", 1) if "!" in addr else (None, addr)
        a = _find_action_at(actions, sheet, cell)
        if not a:
            failures.append(f"no write to {addr}")
            continue
        p = a.get("payload") or {}
        formula = str(p.get("formula") or p.get("value") or "")
        for sub in required_substrings:
            if sub not in formula:
                failures.append(f"{addr} missing '{sub}' (got: {formula!r})")
    if failures:
        return False, "; ".join(failures)
    return True, "formula patterns satisfied"


def formula_at_cell_forbids(actions: list[dict], spec: dict[str, list[str]]) -> tuple[bool, str]:
    failures = []
    for addr, forbidden in spec.items():
        sheet, cell = addr.split("!", 1) if "!" in addr else (None, addr)
        a = _find_action_at(actions, sheet, cell)
        if not a:
            continue  # nothing written, nothing forbidden
        p = a.get("payload") or {}
        formula = str(p.get("formula") or p.get("value") or "")
        for bad in forbidden:
            if bad in formula:
                failures.append(f"{addr} contains forbidden pattern '{bad}' (got: {formula!r})")
    if failures:
        return False, "; ".join(failures)
    return True, "no forbidden patterns in formulas"


def forbidden_sheet_targets(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    forbidden = {s.casefold() for s in spec}
    referenced = _iter_action_sheets(actions)
    hits = [s for s in referenced if s.casefold() in forbidden]
    if hits:
        return False, f"actions reference forbidden sheets: {sorted(set(hits))}"
    return True, f"no actions target forbidden sheets"


def allowed_sheet_targets(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    allowed = {s.casefold() for s in spec}
    referenced = _iter_action_sheets(actions)
    hits = [s for s in referenced if s.casefold() not in allowed]
    if hits:
        return False, f"actions target sheets outside allowed list: {sorted(set(hits))} (allowed: {spec})"
    return True, f"all sheet targets within allowed list"


def must_not_overwrite(actions: list[dict], spec: list[str]) -> tuple[bool, str]:
    """Assert no write/format action targets any cell inside the forbidden ranges."""
    failures = []
    for a in actions:
        t = a.get("type", "")
        if t not in ("write_cell", "write_formula", "write_range", "clear_range"):
            continue
        p = a.get("payload") or {}
        s = p.get("sheet")
        target = p.get("cell") or p.get("address") or p.get("range")
        if not target:
            continue
        addr = _cell_addr(s, target)
        for forbidden_range in spec:
            # Each cell of target range must not fall inside forbidden_range
            if _cell_in_range(addr, forbidden_range) or _cell_in_range(forbidden_range, addr):
                failures.append(f"{t} at {addr} overlaps forbidden range {forbidden_range}")
                break
    if failures:
        return False, "; ".join(failures)
    return True, "no forbidden overwrites"


def min_action_count(actions: list[dict], spec: int) -> tuple[bool, str]:
    if len(actions) < spec:
        return False, f"got {len(actions)} actions, need ≥ {spec}"
    return True, f"{len(actions)} actions (≥ {spec})"


def max_action_count(actions: list[dict], spec: int) -> tuple[bool, str]:
    """Cap on emitted actions. Use for: discuss-mode (max=0), efficient
    formula-fill patterns (max=5 for write+fill_down rather than 38 writes)."""
    if len(actions) > spec:
        return False, f"got {len(actions)} actions, max allowed is {spec}"
    return True, f"{len(actions)} actions (≤ {spec})"


def must_include_action_types_any_of(actions: list[dict],
                                       spec: list[str]) -> tuple[bool, str]:
    """At least ONE of the listed action types must appear. Use when there
    are multiple legitimate ways to satisfy the request (e.g. either
    `add_conditional_format` or the new `conditional_format_advanced`)."""
    types = {a.get("type") for a in actions}
    matched = [t for t in spec if t in types]
    if matched:
        return True, f"matched: {matched}"
    return False, f"none of {spec} present in emitted types {sorted(types)}"


def reply_not_empty(actions: list[dict], spec: bool, reply: str = "",
                    cu_session_id: str | None = None) -> tuple[bool, str]:
    """Sanity: reply text should exist unless a CU session took over.

    Backend intentionally blanks the reply when cu_actions are emitted (the
    add-in's typing animation covers UX while the desktop agent works). So
    "empty reply AND has cu_session_id" is valid — the check only fails when
    BOTH are absent.
    """
    has_reply = bool((reply or "").strip())
    has_cu = bool(cu_session_id)
    ok = has_reply or has_cu
    if spec and not ok:
        return False, "reply is empty and no cu_session_id"
    if has_cu and not has_reply:
        return True, f"reply blanked (cu_session_id={cu_session_id})"
    return True, "reply present"


def cu_session_id_required(actions: list[dict], spec: bool, reply: str = "",
                            cu_session_id: str | None = None) -> tuple[bool, str]:
    """Assert that a CU session was created (i.e. at least one action was
    routed to the desktop agent path). Use this for cases where the action
    type is in COMPUTER_USE_ACTIONS (goal_seek, run_solver, smartart_diagram,
    pivot_table, etc.) — those actions don't appear in `data.actions`, so
    `must_include_action_types` can't see them. Instead, the presence of a
    cu_session_id confirms the model emitted SOMETHING that needs the agent.
    """
    if spec and not cu_session_id:
        return False, "expected a cu_session_id but none was returned"
    if cu_session_id:
        return True, f"cu_session_id present: {cu_session_id}"
    return True, "no cu_session_id required"


# ── Dispatcher ───────────────────────────────────────────────────────────────

CHECKS: dict[str, Callable[..., tuple[bool, str]]] = {
    "must_include_action_types":         must_include_action_types,
    "must_include_action_types_any_of":  must_include_action_types_any_of,
    "must_reference_cells":              must_reference_cells,
    "formula_at_cell_contains":          formula_at_cell_contains,
    "formula_at_cell_forbids":           formula_at_cell_forbids,
    "forbidden_sheet_targets":           forbidden_sheet_targets,
    "allowed_sheet_targets":             allowed_sheet_targets,
    "must_not_overwrite":                must_not_overwrite,
    "min_action_count":                  min_action_count,
    "max_action_count":                  max_action_count,
    "reply_not_empty":                   reply_not_empty,
    "cu_session_id_required":            cu_session_id_required,
}


def evaluate(rubric: dict, actions: list[dict], reply: str = "",
             cu_session_id: str | None = None) -> list[tuple[bool, str, str]]:
    """Run every check listed in `rubric`. Returns list of (pass, name, detail)."""
    results: list[tuple[bool, str, str]] = []
    for name, spec in rubric.items():
        if name.startswith("_"):  # comments / metadata
            continue
        check = CHECKS.get(name)
        if check is None:
            results.append((False, name, f"UNKNOWN RUBRIC KEY"))
            continue
        try:
            if name in ("reply_not_empty", "cu_session_id_required"):
                ok, detail = check(actions, spec, reply=reply, cu_session_id=cu_session_id)
            else:
                ok, detail = check(actions, spec)
        except Exception as e:
            ok, detail = False, f"check raised: {type(e).__name__}: {e}"
        results.append((ok, name, detail))
    return results
