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

from rubrics import excel as excel_rubric  # noqa: E402

DEFAULT_BACKEND = "https://focused-solace-production-6839.up.railway.app"
CASES_DIR = HERE / "cases"
REPORT_DIR = HERE / "_reports"

APP_RUBRICS = {
    "excel": excel_rubric.evaluate,
    # "word":       word_rubric.evaluate,
    # "powerpoint": powerpoint_rubric.evaluate,
    # "rstudio":    rstudio_rubric.evaluate,
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

def run_case(case: dict, backend: str, timeout: int) -> dict:
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

    check_results = evaluator(rubric, actions, reply)
    all_ok = all(ok for ok, _, _ in check_results)

    return {
        "name": name, "ok": all_ok, "error": None,
        "checks": [{"pass": ok, "name": n, "detail": d} for ok, n, d in check_results],
        "elapsed_ms": elapsed_ms,
        "actions": actions, "reply": reply,
        "cu_session_id": data.get("cu_session_id"),
    }


def print_case_result(result: dict):
    name = result["name"]
    if result.get("error"):
        print(f"{c_red('✗')} {c_bold(name)} — {c_red('error')}: {result['error']}")
        return

    status = c_green("PASS") if result["ok"] else c_red("FAIL")
    meta = c_dim(f"({result['elapsed_ms']}ms, {len(result['actions'])} actions)")
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
        r = run_case(case, args.backend, args.timeout)
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
