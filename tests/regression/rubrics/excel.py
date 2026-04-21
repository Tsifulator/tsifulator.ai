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


def reply_not_empty(actions: list[dict], spec: bool, reply: str = "") -> tuple[bool, str]:
    """Sanity: reply text should exist (catches empty-reply regressions)."""
    ok = bool(reply.strip())
    if spec and not ok:
        return False, "reply text is empty"
    return True, "reply present" if ok else "no reply required"


# ── Dispatcher ───────────────────────────────────────────────────────────────

CHECKS: dict[str, Callable[..., tuple[bool, str]]] = {
    "must_include_action_types": must_include_action_types,
    "must_reference_cells":      must_reference_cells,
    "formula_at_cell_contains":  formula_at_cell_contains,
    "formula_at_cell_forbids":   formula_at_cell_forbids,
    "forbidden_sheet_targets":   forbidden_sheet_targets,
    "allowed_sheet_targets":     allowed_sheet_targets,
    "must_not_overwrite":        must_not_overwrite,
    "min_action_count":          min_action_count,
    "reply_not_empty":           reply_not_empty,
}


def evaluate(rubric: dict, actions: list[dict], reply: str = "") -> list[tuple[bool, str, str]]:
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
            if name == "reply_not_empty":
                ok, detail = check(actions, spec, reply=reply)
            else:
                ok, detail = check(actions, spec)
        except Exception as e:
            ok, detail = False, f"check raised: {type(e).__name__}: {e}"
        results.append((ok, name, detail))
    return results
