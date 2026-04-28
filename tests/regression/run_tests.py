#!/usr/bin/env python3
"""tsifl regression harness — runs every case in tests/regression/cases/ against
a tsifl backend (deployed Railway by default) and validates the returned actions
against each case's rubric.yaml.

Exit code 0 = all pass. Non-zero = at least one case failed.

Usage:
    python run_tests.py                              # all cases, Railway backend
    python run_tests.py --case 001-placerhills-09    # one case
    python run_tests.py --backend http://localhost:8000
    python run_tests.py --json                       # machine-readable output
"""
from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import sys
import time
from pathlib import Path

import requests
import yaml

# Ensure rubrics package importable from this directory
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from rubrics import excel   as excel_rubric    # noqa: E402
from rubrics import rstudio as rstudio_rubric  # noqa: E402

DEFAULT_BACKEND = "https://focused-solace-production-6839.up.railway.app"
CASES_DIR = HERE / "cases"
REPORT_DIR = HERE / "_reports"

APP_RUBRICS = {
    "excel":   excel_rubric.evaluate,
    "rstudio": rstudio_rubric.evaluate,
    # "word":       word_rubric.evaluate,
    # "powerpoint": powerpoint_rubric.evaluate,
}


# ── ANSI colors (no external dep) ───────────────────────────────────────────

def c_green(s): return f"\033[92m{s}\033[0m"
def c_red(s):   return f"\033[91m{s}\033[0m"
def c_dim(s):   return f"\033[90m{s}\033[0m"
def c_bold(s):  return f"\033[1m{s}\033[0m"


# ── Case loading ────────────────────────────────────────────────────────────

def load_case(case_dir: Path) -> dict:
    req_path = case_dir / "request.json"
    if not req_path.exists():
        raise FileNotFoundError(f"{case_dir.name}: missing request.json")
    req = json.loads(req_path.read_text())

    # Attach images from images/ if present
    images_dir = case_dir / "images"
    if images_dir.exists():
        images = []
        for img_path in sorted(images_dir.iterdir()):
            if img_path.is_file() and img_path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}:
                mime, _ = mimetypes.guess_type(str(img_path))
                images.append({
                    "file_name": img_path.name,
                    "media_type": mime or "image/png",
                    "data": base64.b64encode(img_path.read_bytes()).decode("ascii"),
                })
        if images:
            req["images"] = images

    rubric_path = case_dir / "rubric.yaml"
    rubric = yaml.safe_load(rubric_path.read_text()) if rubric_path.exists() else {}

    return {"name": case_dir.name, "request": req, "rubric": rubric}


# ── Runner ──────────────────────────────────────────────────────────────────

def run_case(case: dict, backend: str, timeout: int,
             cheap: bool = False, cached: bool = False,
             refresh_golden: bool = False,
             case_dir: Path | None = None) -> dict:
    name = case["name"]
    req = case["request"]
    rubric = case["rubric"]

    app = (req.get("context") or {}).get("app") or "excel"
    evaluator = APP_RUBRICS.get(app)
    if evaluator is None:
        return {
            "name": name, "ok": False,
            "error": f"no rubric evaluator for app={app!r}",
            "checks": [], "elapsed_ms": 0, "actions": [], "reply": "",
        }

    # --cached: skip the live call entirely and replay the stored response.
    # Useful when iterating on rubric assertions — the evaluator logic runs
    # fully, the LLM does not. Captured by capture_case.py on the first run.
    if cached and case_dir is not None:
        resp_path = case_dir / "response.json"
        if resp_path.exists():
            data = json.loads(resp_path.read_text())
            actions = data.get("actions") or []
            reply = data.get("reply") or ""
            check_results = evaluator(rubric, actions, reply, cu_session_id=data.get("cu_session_id"))
            all_ok = all(ok for ok, _, _ in check_results)
            return {
                "name": name, "ok": all_ok, "error": None,
                "checks": [{"pass": ok, "name": n, "detail": d} for ok, n, d in check_results],
                "elapsed_ms": 0, "actions": actions, "reply": reply,
                "cu_session_id": data.get("cu_session_id"),
                "cached": True,
            }
        # Fall through to live call if no cached response exists yet

    # --cheap: force every case to use MODEL_FAST (Haiku). Cuts regression
    # suite spend ~10×. Keep the pipeline (guards, routing, memory) exercised
    # without burning full-model tokens on every CI run.
    if cheap:
        req = dict(req)
        req["context"] = dict(req.get("context") or {})
        req["context"]["force_model"] = "haiku"

    # Pre-flight: clear any project_memory state for this test user_id so the
    # case starts from a deterministic blank slate. Without this, the first
    # case to run pollutes state for the next. Silently ignores 404s — if the
    # endpoint doesn't exist yet or state is already empty, no harm done.
    #
    # Mirror backend services/project_memory.fingerprint(): prefer
    # `workbook_name` when present (current clients), fall back to the
    # sheets hash. Compute BOTH and clear both — older runs may have
    # written under the legacy hash, current ones write under the
    # workbook-name hash; we want neither to leak between cases.
    user_id = req.get("user_id", "")
    app_ctx = (req.get("context") or {})
    import hashlib as _h, json as _j
    app = app_ctx.get("app") or ""
    name = (app_ctx.get("workbook_name") or "").strip()
    primary_sig = (
        {"app": app, "workbook": name}
        if name else
        {"app": app, "sheets": sorted(app_ctx.get("all_sheets") or [])}
    )
    legacy_sig = {"app": app, "sheets": sorted(app_ctx.get("all_sheets") or [])}
    workbook_ids = {
        _h.sha256(_j.dumps(primary_sig, sort_keys=True).encode()).hexdigest()[:16],
        _h.sha256(_j.dumps(legacy_sig,  sort_keys=True).encode()).hexdigest()[:16],
    }
    workbook_id = next(iter(workbook_ids))  # for any later debug log
    for wid in workbook_ids:
        try:
            requests.delete(
                f"{backend.rstrip('/')}/chat/debug/project-memory/{user_id}/{wid}",
                timeout=10,
            )
        except Exception:
            pass  # pre-flight is best-effort

    t0 = time.time()
    try:
        resp = requests.post(f"{backend.rstrip('/')}/chat/", json=req, timeout=timeout)
    except requests.RequestException as e:
        return {
            "name": name, "ok": False,
            "error": f"request failed: {type(e).__name__}: {e}",
            "checks": [], "elapsed_ms": int((time.time() - t0) * 1000),
            "actions": [], "reply": "",
        }
    elapsed_ms = int((time.time() - t0) * 1000)

    if resp.status_code != 200:
        return {
            "name": name, "ok": False,
            "error": f"HTTP {resp.status_code}: {resp.text[:300]}",
            "checks": [], "elapsed_ms": elapsed_ms,
            "actions": [], "reply": "",
        }

    data = resp.json()
    actions = data.get("actions") or []
    reply = data.get("reply") or ""
    cu_session_id = data.get("cu_session_id")

    # response.json is the "golden" response used by --cached. Default live
    # runs MUST NOT touch it (flaky LLM output could silently poison the
    # golden). Only overwrite when --refresh-golden is explicitly set.
    if refresh_golden and case_dir is not None:
        try:
            (case_dir / "response.json").write_text(json.dumps(data, indent=2))
        except Exception:
            pass

    check_results = evaluator(rubric, actions, reply, cu_session_id=cu_session_id)
    all_ok = all(ok for ok, _, _ in check_results)

    return {
        "name": name, "ok": all_ok, "error": None,
        "checks": [{"pass": ok, "name": n, "detail": d} for ok, n, d in check_results],
        "elapsed_ms": elapsed_ms,
        "actions": actions, "reply": reply,
        "cu_session_id": data.get("cu_session_id"),
        "cached": False,
    }


def print_case_result(result: dict):
    name = result["name"]
    if result.get("error"):
        print(f"{c_red('✗')} {c_bold(name)} — {c_red('error')}: {result['error']}")
        return

    status = c_green("PASS") if result["ok"] else c_red("FAIL")
    source = "cached" if result.get("cached") else f"{result['elapsed_ms']}ms"
    meta = c_dim(f"({source}, {len(result['actions'])} actions)")
    print(f"  {status}  {c_bold(name)}  {meta}")
    for chk in result["checks"]:
        mark = c_green("✓") if chk["pass"] else c_red("✗")
        print(f"    {mark} {chk['name']}  {c_dim('— ' + chk['detail'])}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", help="run a single case (directory name)")
    ap.add_argument("--backend", default=os.environ.get("TSIFL_BACKEND") or DEFAULT_BACKEND)
    ap.add_argument("--timeout", type=int, default=120)
    ap.add_argument("--json", action="store_true", help="emit JSON-only output")
    ap.add_argument("--save-report", action="store_true", help="write full report JSON to _reports/")
    ap.add_argument("--cheap", action="store_true",
                    help="Force Haiku on every case. ~10x cheaper than default Sonnet. "
                         "Exercises guards + routing + memory without burning full-model tokens. "
                         "Still makes a live API call — use --cached to skip API entirely.")
    ap.add_argument("--cached", action="store_true",
                    help="Skip the live /chat call and replay each case's saved response.json. "
                         "Falls back to live if no response is saved. $0 cost; use when iterating "
                         "on rubric assertions, not when validating backend changes.")
    ap.add_argument("--refresh-golden", action="store_true",
                    help="DESTRUCTIVE: after each live run, overwrite the case's response.json "
                         "with the fresh response. Use ONLY when you intentionally want to update "
                         "the golden (e.g. LLM behavior drifted legitimately). Default behavior "
                         "never touches response.json.")
    args = ap.parse_args()

    if not CASES_DIR.exists():
        print(f"No cases directory at {CASES_DIR}", file=sys.stderr)
        return 2

    dirs = sorted(d for d in CASES_DIR.iterdir() if d.is_dir())
    if args.case:
        dirs = [d for d in dirs if d.name == args.case]
        if not dirs:
            print(f"Case not found: {args.case}", file=sys.stderr)
            return 2

    if not args.json:
        print(c_bold(f"\ntsifl regression suite — backend={args.backend}"))
        print(c_dim(f"running {len(dirs)} case(s)\n"))

    results = []
    for d in dirs:
        try:
            case = load_case(d)
        except Exception as e:
            results.append({"name": d.name, "ok": False, "error": f"load failed: {e}",
                            "checks": [], "elapsed_ms": 0, "actions": [], "reply": ""})
            continue
        r = run_case(case, args.backend, args.timeout,
                     cheap=args.cheap, cached=args.cached,
                     refresh_golden=args.refresh_golden, case_dir=d)
        results.append(r)
        if not args.json:
            print_case_result(r)

    passed = sum(1 for r in results if r["ok"])
    failed = len(results) - passed

    if args.json:
        print(json.dumps({"passed": passed, "failed": failed, "results": results}, indent=2))
    else:
        summary = c_green(f"{passed} passed") if passed else c_dim("0 passed")
        if failed:
            summary += "  " + c_red(f"{failed} failed")
        print(f"\n{summary}\n")

    if args.save_report:
        REPORT_DIR.mkdir(exist_ok=True)
        stamp = time.strftime("%Y%m%d-%H%M%S")
        (REPORT_DIR / f"report-{stamp}.json").write_text(
            json.dumps({"backend": args.backend, "results": results}, indent=2)
        )

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
