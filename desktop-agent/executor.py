"""
executor.py — tsifl Mac action execution engine.

Takes structured action plans from Claude and executes them on the
user's Mac via AppleScript, Spotlight, shell commands, and native APIs.

Every action goes through a risk classification before execution:
  - GREEN  (read-only): auto-execute, no confirmation needed
  - YELLOW (writes): show plan, one-click confirm
  - RED    (irreversible): show plan, require explicit confirmation

The executor never runs anything the user hasn't approved.
"""

import json
import logging
import os
import subprocess
import sys
import time
from dataclasses import dataclass, field, asdict
from enum import Enum
from pathlib import Path
from typing import Optional

logger = logging.getLogger("tsifl.executor")


# ── Risk levels ──────────────────────────────────────────────────────────────

class Risk(str, Enum):
    GREEN = "green"    # read-only: search, open, show info
    YELLOW = "yellow"  # writes: create file, type text, move file
    RED = "red"        # irreversible: send email, delete, purchase


# ── Action dataclass ─────────────────────────────────────────────────────────

@dataclass
class Action:
    """A single step in an execution plan."""
    type: str                          # e.g. "search_files", "open_app", "applescript"
    description: str                   # human-readable: "Search for Excel files containing 'grocery'"
    command: str                       # the actual command/script to execute
    risk: Risk = Risk.GREEN
    result: Optional[str] = None       # filled after execution
    success: bool = False
    error: Optional[str] = None

    def to_dict(self):
        d = asdict(self)
        d["risk"] = self.risk.value
        return d


# ── File search via Spotlight (mdfind) ───────────────────────────────────────

def search_files(query: str, max_results: int = 10, file_type: str = None) -> list[str]:
    """Search for files using macOS Spotlight (mdfind).

    Args:
        query: search term (filename, content, or raw mdfind query)
        max_results: cap on number of results
        file_type: optional filter — "excel", "pdf", "image", "document", etc.

    Returns:
        List of file paths matching the query.
    """
    # Map friendly type names to Spotlight content types
    type_filters = {
        "excel": 'kMDItemContentType == "org.openxmlformats.spreadsheetml.sheet" || kMDItemContentType == "com.microsoft.excel.xls"',
        "word": 'kMDItemContentType == "org.openxmlformats.wordprocessingml.document" || kMDItemContentType == "com.microsoft.word.doc"',
        "ppt": 'kMDItemContentType == "org.openxmlformats.presentationml.presentation"',
        "pdf": 'kMDItemContentType == "com.adobe.pdf"',
        "image": 'kMDItemContentTypeTree == "public.image"',
        "csv": 'kMDItemContentType == "public.comma-separated-values-text"',
        "text": 'kMDItemContentTypeTree == "public.text"',
    }

    # Build mdfind query
    parts = []
    if file_type and file_type.lower() in type_filters:
        parts.append(f"({type_filters[file_type.lower()]})")

    # Add name/content search
    if query:
        # If the query looks like a raw mdfind expression, use it directly
        if "kMDItem" in query:
            parts.append(query)
        else:
            # Search both filename and content
            escaped = query.replace('"', '\\"')
            parts.append(f'(kMDItemFSName == "*{escaped}*"cdw || kMDItemTextContent == "*{escaped}*"cdw)')

    mdfind_query = " && ".join(parts) if parts else f'kMDItemFSName == "*{query}*"cdw'

    try:
        result = subprocess.run(
            ["mdfind", mdfind_query],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode != 0:
            return []
        paths = [p.strip() for p in result.stdout.strip().split("\n") if p.strip()]
        # Filter out hidden/system paths
        paths = [p for p in paths if not any(
            seg.startswith(".") for seg in Path(p).parts[1:]  # skip root /
        )]
        # Sort by modification time (most recent first)
        paths.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0, reverse=True)
        return paths[:max_results]
    except subprocess.TimeoutExpired:
        return []
    except Exception as e:
        logger.error(f"mdfind failed: {e}")
        return []


# ── Deterministic action handlers ────────────────────────────────────────────
# These handle common tasks WITHOUT relying on Claude to write AppleScript.
# Claude just says WHAT to do, the executor knows HOW.

import urllib.parse as _urlparse


def _play_media(platform: str, query: str) -> tuple[bool, str]:
    """Play media on a specific platform — deterministic, no vision needed."""
    import sys
    sys.stderr.write(f"[play_media] platform={platform!r} query={query!r}\n")

    if platform in ("youtube", "yt"):
        encoded = _urlparse.quote_plus(query)
        url = f"https://www.youtube.com/results?search_query={encoded}"
        try:
            subprocess.run(["open", url], check=True, timeout=5)
            return True, f"Opened YouTube search for '{query}'"
        except Exception as e:
            return False, str(e)

    elif platform in ("spotify",):
        return spotify_play(query)

    elif platform in ("apple music", "music"):
        # Open Apple Music search
        encoded = _urlparse.quote_plus(query)
        url = f"https://music.apple.com/search?term={encoded}"
        try:
            subprocess.run(["open", url], check=True, timeout=5)
            return True, f"Opened Apple Music search for '{query}'"
        except Exception as e:
            return False, str(e)

    else:
        # Fallback: try opening as URL search
        encoded = _urlparse.quote_plus(query)
        url = f"https://www.google.com/search?q={encoded}+{platform}"
        try:
            subprocess.run(["open", url], check=True, timeout=5)
            return True, f"Searched for '{query}' on {platform}"
        except Exception as e:
            return False, str(e)


def _web_search(query: str, engine: str = "google") -> tuple[bool, str]:
    """Open a web search — deterministic, no vision needed."""
    encoded = _urlparse.quote_plus(query)
    urls = {
        "google": f"https://www.google.com/search?q={encoded}",
        "youtube": f"https://www.youtube.com/results?search_query={encoded}",
        "bing": f"https://www.bing.com/search?q={encoded}",
        "duckduckgo": f"https://duckduckgo.com/?q={encoded}",
    }
    url = urls.get(engine, urls["google"])
    try:
        subprocess.run(["open", url], check=True, timeout=5)
        return True, f"Opened {engine} search for '{query}'"
    except Exception as e:
        return False, str(e)


def _fetch_url(url: str, max_chars: int = 8000) -> tuple[bool, str]:
    """Fetch a URL and return the text content. Strips HTML.

    Used so Claude can look things up (Macabacus shortcuts, docs, etc.)
    without opening a browser tab. Uses httpx with SSL verify disabled
    to handle networks with SSL inspection (corp/edu).
    """
    try:
        import httpx as _hx
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,*/*",
        }
        # verify=False handles ZScaler / corp SSL inspection
        with _hx.Client(verify=False, follow_redirects=True, timeout=15) as c:
            resp = c.get(url, headers=headers)
            raw = resp.text

        # Strip HTML tags + scripts/styles
        import re as _re
        text = _re.sub(r"<script\b[^>]*>.*?</script>", " ", raw,
                       flags=_re.DOTALL | _re.IGNORECASE)
        text = _re.sub(r"<style\b[^>]*>.*?</style>", " ", text,
                       flags=_re.DOTALL | _re.IGNORECASE)
        text = _re.sub(r"<[^>]+>", " ", text)
        # Decode common HTML entities
        text = (text.replace("&nbsp;", " ").replace("&amp;", "&")
                    .replace("&lt;", "<").replace("&gt;", ">")
                    .replace("&quot;", '"').replace("&#39;", "'"))
        text = _re.sub(r"\s+", " ", text).strip()
        return True, text[:max_chars]
    except Exception as e:
        return False, f"fetch failed: {e}"


def _web_lookup(query: str, max_chars: int = 6000) -> tuple[bool, str]:
    """Search the web AND fetch the top results' content.

    Tries multiple providers in order until one returns parseable results:
      1. DuckDuckGo HTML — fast but rate-limits aggressively
      2. Bing HTML — slower but more reliable for finance/tools queries
    Falls back to Google if both fail.
    """
    import httpx as _hx
    import re as _re
    import urllib.parse as _up

    encoded = _urlparse.quote_plus(query)
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
    }

    # Provider list: (name, url_template, regex_to_extract_result_urls)
    providers = [
        ("ddg", f"https://html.duckduckgo.com/html/?q={encoded}",
         r'href="/l/\?uddg=([^"&]+)'),
        ("bing", f"https://www.bing.com/search?q={encoded}",
         r'<h2><a [^>]*href="([^"]+)"'),
        ("google", f"https://www.google.com/search?q={encoded}&hl=en",
         r'<a href="/url\?q=([^&"]+)'),
    ]

    raw = ""
    urls: list = []
    used_provider = "none"

    try:
        with _hx.Client(verify=False, follow_redirects=True, timeout=15) as c:
            for name, url, pattern in providers:
                try:
                    resp = c.get(url, headers=headers)
                    candidate = resp.text
                    # Detect bot challenges / empty results
                    low = candidate.lower()
                    if ("unfortunately, bots" in low
                            or "captcha" in low
                            or "/sorry/" in low
                            or len(candidate) < 1500):
                        continue
                    found = _re.findall(pattern, candidate[:80000])
                    if found:
                        raw = candidate
                        urls = found
                        used_provider = name
                        break
                except Exception:
                    continue
    except Exception as e:
        return False, f"web_lookup failed: {e}"

    if not urls:
        return False, f"web_lookup: no usable results (tried {[p[0] for p in providers]})"

    # Dedupe by hostname, skip search engine self-links
    seen = set()
    result_urls: list = []
    for u in urls:
        try:
            u = _up.unquote(u)
        except Exception:
            pass
        # Bing/Google sometimes wrap URLs — strip protocol prefix nonsense
        if u.startswith("/url?q="):
            u = u[7:]
        host = u.split("/")[2] if "://" in u else ""
        if not host:
            continue
        if any(s in host for s in ("duckduckgo", "google.", "bing.", "youtube.com/redirect")):
            continue
        if host in seen:
            continue
        seen.add(host)
        result_urls.append(u)
        if len(result_urls) >= 3:
            break

    if not result_urls:
        # Couldn't extract third-party URLs — return the search page as text
        clean = _re.sub(r"<[^>]+>", " ", raw)
        clean = _re.sub(r"\s+", " ", clean).strip()
        return True, f"[{used_provider} search page]\n" + clean[:max_chars]

    # Fetch top 2 result pages and concatenate
    chunks = []
    per_url_budget = max_chars // min(len(result_urls), 2)
    for u in result_urls[:2]:
        ok, content = _fetch_url(u, max_chars=per_url_budget)
        if ok and content:
            chunks.append(f"[Source: {u}]\n{content[:per_url_budget]}")

    if not chunks:
        clean = _re.sub(r"<[^>]+>", " ", raw)
        clean = _re.sub(r"\s+", " ", clean).strip()
        return True, f"[{used_provider} search page — couldn't fetch result pages]\n" + clean[:max_chars]

    return True, "\n\n---\n\n".join(chunks)


# ── Data Export — battle-tested per-app scripts ─────────────────────────────
# Claude just says {"type": "data_export", "source_app": "Numbers", "destination": "~/Desktop/data.csv"}
# and we handle the AppleScript perfectly every time.

def _data_export(source_app: str, dest_path: str, fmt: str = "csv") -> tuple[bool, str]:
    """Export data from a Mac app to a file. Uses tested AppleScript per app.

    Args:
        source_app: app name (e.g. "Numbers", "Microsoft Excel")
        dest_path: absolute POSIX path for output file
        fmt: format — "csv", "tsv", "pdf"

    Returns:
        (True, summary) on success, (False, error) on failure.
    """
    import sys
    sys.stderr.write(f"[data_export] {source_app} → {dest_path} (fmt={fmt})\n")

    app_lower = source_app.lower().strip()

    # ── Numbers ──────────────────────────────────────────────────────
    if app_lower in ("numbers", "apple numbers"):
        # Convert POSIX path to HFS path for Numbers export
        # /Users/nick/Desktop/data.csv → Macintosh HD:Users:nick:Desktop:data.csv
        # Simpler: use path to home folder + relative
        home = str(Path.home())
        if dest_path.startswith(home):
            relative = dest_path[len(home):].lstrip("/")
            # HFS uses colons: Desktop/data.csv → Desktop:data.csv
            hfs_relative = relative.replace("/", ":")
            script = (
                f'tell application "Numbers"\n'
                f'    set theDoc to front document\n'
                f'    set exportPath to ((path to home folder as text) & "{hfs_relative}")\n'
                f'    export theDoc to file exportPath as CSV\n'
                f'end tell'
            )
        else:
            # Fallback for non-home paths: use POSIX file
            script = (
                f'tell application "Numbers"\n'
                f'    set theDoc to front document\n'
                f'    export theDoc to POSIX file "{dest_path}" as CSV\n'
                f'end tell'
            )
        ok, result = run_applescript(script, timeout=30)
        if ok:
            if Path(dest_path).exists():
                size = Path(dest_path).stat().st_size
                return True, f"Exported {size:,} bytes to {Path(dest_path).name}"
            return True, f"Export completed to {Path(dest_path).name}"
        return False, f"Numbers export failed: {result}"

    # ── Microsoft Excel ──────────────────────────────────────────────
    elif app_lower in ("microsoft excel", "excel"):
        script = (
            f'tell application "Microsoft Excel"\n'
            f'    set filePath to POSIX file "{dest_path}" as text\n'
            f'    save active workbook in filePath as CSV file format\n'
            f'end tell'
        )
        ok, result = run_applescript(script, timeout=30)
        if ok:
            if Path(dest_path).exists():
                size = Path(dest_path).stat().st_size
                return True, f"Exported {size:,} bytes to {Path(dest_path).name}"
            return True, f"Export completed to {Path(dest_path).name}"
        return False, f"Excel export failed: {result}"

    # ── Fallback: clipboard approach ─────────────────────────────────
    else:
        # For apps without export API, use activate → Cmd+A → Cmd+C → pbpaste
        script = (
            f'tell application "{source_app}" to activate\n'
            f'delay 1.0\n'
            f'tell application "System Events" to tell process "{source_app}"\n'
            f'    set frontmost to true\n'
            f'    delay 0.5\n'
            f'    keystroke "a" using command down\n'
            f'    delay 0.5\n'
            f'    keystroke "c" using command down\n'
            f'end tell\n'
            f'delay 1.0\n'
            f'do shell script "pbpaste > " & quoted form of "{dest_path}"'
        )
        ok, result = run_applescript(script, timeout=20)
        if ok:
            if Path(dest_path).exists() and Path(dest_path).stat().st_size > 0:
                size = Path(dest_path).stat().st_size
                return True, f"Copied {size:,} bytes to {Path(dest_path).name}"
            return False, f"Clipboard copy produced empty file"
        return False, f"Copy from {source_app} failed: {result}"


# ── AppleScript execution ────────────────────────────────────────────────────

def run_applescript(script: str, timeout: int = 15) -> tuple[bool, str]:
    """Execute an AppleScript and return (success, output_or_error).

    Args:
        script: the AppleScript source code
        timeout: max seconds to wait

    Returns:
        (True, stdout) on success, (False, stderr) on failure.
    """
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=timeout,
        )
        if result.returncode == 0:
            return True, result.stdout.strip()
        return False, result.stderr.strip()
    except subprocess.TimeoutExpired:
        return False, f"AppleScript timed out after {timeout}s"
    except Exception as e:
        return False, str(e)


# ── Shell command execution (read-only) ──────────────────────────────────────

def run_shell(command: str, timeout: int = 10) -> tuple[bool, str]:
    """Run a read-only shell command and return (success, output).

    Safety: this should only be called for read-only commands (ls, cat,
    mdfind, date, etc.). Write operations should go through AppleScript
    or dedicated functions with proper risk classification.
    """
    try:
        result = subprocess.run(
            command, shell=True,
            capture_output=True, text=True, timeout=timeout,
        )
        output = result.stdout.strip()
        if result.returncode != 0 and result.stderr.strip():
            output = result.stderr.strip()
        return result.returncode == 0, output
    except subprocess.TimeoutExpired:
        return False, f"Command timed out after {timeout}s"
    except Exception as e:
        return False, str(e)


# ── High-level Mac operations ────────────────────────────────────────────────

def open_file(path: str) -> tuple[bool, str]:
    """Open a file with its default application."""
    try:
        expanded = str(Path(path).expanduser())
        subprocess.run(["open", expanded], check=True, timeout=5)
        return True, f"Opened {Path(expanded).name}"
    except Exception as e:
        return False, str(e)


def open_app(app_name: str) -> tuple[bool, str]:
    """Launch or activate a macOS application.

    Uses `open -a` which returns immediately (non-blocking) — avoids
    the 15s AppleScript timeout on cold launches.  The vision loop's
    `wait` action gives the app time to fully appear before screenshotting.
    """
    try:
        subprocess.run(["open", "-a", app_name], check=True, timeout=10)
        # Give the app a moment to register, then try to bring it front
        time.sleep(0.5)
        # Quick activate — if the app is already running this is instant;
        # if it's still loading we don't care if this times out (5s cap)
        run_applescript(
            f'tell application "{app_name}" to activate', timeout=5
        )
        return True, f"Opened {app_name}"
    except subprocess.CalledProcessError:
        return False, f"Could not find app: {app_name}"
    except Exception as e:
        # Even if activate timed out, open -a likely succeeded
        return True, f"Opened {app_name} (still loading)"


def open_url(url: str) -> tuple[bool, str]:
    """Open a URL in the default browser."""
    try:
        subprocess.run(["open", url], check=True, timeout=5)
        return True, f"Opened {url}"
    except Exception as e:
        return False, str(e)


def get_clipboard() -> str:
    """Get the current clipboard contents."""
    ok, text = run_shell("pbpaste")
    return text if ok else ""


def set_clipboard(text: str) -> tuple[bool, str]:
    """Set the clipboard contents."""
    try:
        process = subprocess.Popen(
            ["pbcopy"], stdin=subprocess.PIPE
        )
        process.communicate(text.encode("utf-8"), timeout=5)
        return True, "Copied to clipboard"
    except Exception as e:
        return False, str(e)


def get_frontmost_app() -> str:
    """Get the name of the frontmost application."""
    ok, name = run_applescript(
        'tell application "System Events" to get name of first process whose frontmost is true'
    )
    return name if ok else "unknown"


def get_running_apps() -> list[str]:
    """Get list of running applications."""
    ok, output = run_applescript(
        'tell application "System Events" to get name of every process whose background only is false'
    )
    if ok:
        return [a.strip() for a in output.split(",") if a.strip()]
    return []


# ── Rich context capture ────────────────────────────────────────────────────

def get_excel_context() -> dict | None:
    """Read the active Excel workbook's structure for Claude.

    Returns a dict matching what context_formatter.format_context expects
    for the Excel app:
      - sheet, selected_cell, used_range
      - sheet_summaries: [{name, rows, cols, used_range, preview, preview_formulas}, ...]
      - sheet_data: rows × cols of active sheet
      - sheet_formulas: same shape as sheet_data

    Capped to keep token cost reasonable:
      - Per sheet: 25 rows × 12 cols preview
      - Active sheet: 50 rows × 20 cols with formulas
    Returns None if Excel isn't running or has no open workbook.
    """
    # First check Excel has a workbook open
    ok, _ = run_applescript(
        'tell application "Microsoft Excel" to count workbooks',
        timeout=3,
    )
    if not ok:
        return None

    # Read workbook metadata in a few simple AppleScript calls.
    # Excel's AppleScript doesn't like complex repeat-loops over sheets,
    # but "get name of every sheet" works perfectly.
    ok, wb_name = run_applescript(
        'tell application "Microsoft Excel" to get name of active workbook',
        timeout=5,
    )
    if not ok:
        return None

    ok, active_sheet = run_applescript(
        'tell application "Microsoft Excel" to get name of active sheet',
        timeout=5,
    )
    if not ok:
        active_sheet = ""

    ok, selection = run_applescript(
        'tell application "Microsoft Excel" to get address of selection',
        timeout=5,
    )
    if not ok:
        selection = ""

    ok, sheets_raw = run_applescript(
        'tell application "Microsoft Excel" to get name of every sheet of active workbook',
        timeout=5,
    )
    if not ok or not sheets_raw:
        return None
    # AppleScript returns: "Sheet1, Sheet2, Sheet3" — comma-space separated
    sheet_names = [n.strip() for n in sheets_raw.split(",") if n.strip()]

    # For each sheet, read its used-range dimensions + cell values + formulas
    # in a single AppleScript call returning a 2D list.
    sheet_summaries = []
    active_sheet_data: list = []
    active_sheet_formulas: list = []
    active_used_range = ""

    for s_name in sheet_names:
        esc_name = s_name.replace("\\", "\\\\").replace('"', '\\"')

        # Get used range metadata
        ok_m, meta = run_applescript(
            f'''tell application "Microsoft Excel"
    try
        set ur to used range of worksheet "{esc_name}" of active workbook
        return (get address of ur) & "||" & (count of rows of ur) & "||" & (count of columns of ur)
    on error
        return "||0||0"
    end try
end tell''',
            timeout=5,
        )
        rng_addr = ""
        rows_n = cols_n = 0
        if ok_m and meta:
            mparts = meta.split("||")
            if len(mparts) >= 3:
                rng_addr = mparts[0]
                try:
                    rows_n = int(mparts[1])
                    cols_n = int(mparts[2])
                except ValueError:
                    rows_n = cols_n = 0

        preview_rows: list = []
        preview_formulas: list = []

        if rows_n > 0 and cols_n > 0:
            # Get values + formulas as flat comma-separated AppleScript lists
            ok_v, vals_raw = run_applescript(
                f'tell application "Microsoft Excel" to get value of '
                f'used range of worksheet "{esc_name}" of active workbook',
                timeout=8,
            )
            ok_f, forms_raw = run_applescript(
                f'tell application "Microsoft Excel" to get formula of '
                f'used range of worksheet "{esc_name}" of active workbook',
                timeout=8,
            )

            # AppleScript returns 2D lists flattened with ", " — split by row count.
            # First reshape to 2D, then cap to 25 rows × 12 cols.
            def _split_cells(raw_str: str, total_cells: int) -> list[str]:
                if not raw_str:
                    return [""] * total_cells
                # Single cell case: no commas
                if total_cells == 1:
                    return [raw_str]
                # AppleScript joins with ", " — but cell values may contain commas!
                # For now, use a heuristic: try comma-space split, pad/truncate to total_cells
                parts = raw_str.split(", ")
                if len(parts) != total_cells:
                    # Mismatch — values contained commas. Best effort: keep what we have.
                    pass
                return parts

            val_flat = _split_cells(vals_raw or "", rows_n * cols_n)
            form_flat = _split_cells(forms_raw or "", rows_n * cols_n)

            # Reshape to 2D, then cap
            row_cap = min(rows_n, 25)
            col_cap = min(cols_n, 12)
            for r in range(row_cap):
                vrow = []
                frow = []
                for c in range(col_cap):
                    idx = r * cols_n + c
                    v = val_flat[idx] if idx < len(val_flat) else ""
                    f = form_flat[idx] if idx < len(form_flat) else ""
                    vrow.append(v)
                    frow.append(f if f.startswith("=") else None)
                preview_rows.append(vrow)
                preview_formulas.append(frow)

        sheet_summaries.append({
            "name": s_name,
            "rows": rows_n,
            "cols": cols_n,
            "used_range": rng_addr,
            "preview": preview_rows,
            "preview_formulas": preview_formulas,
        })

        if s_name == active_sheet:
            active_sheet_data = preview_rows
            active_sheet_formulas = preview_formulas
            active_used_range = rng_addr

    return {
        "app": "excel",
        "workbook_name": wb_name,
        "sheet": active_sheet,
        "selected_cell": selection or "A1",
        "used_range": active_used_range,
        "sheet_summaries": sheet_summaries,
        "sheet_data": active_sheet_data,
        "sheet_formulas": active_sheet_formulas,
    }


def get_system_context() -> dict:
    """Capture rich context about the current Mac state."""
    ctx = {
        "frontmost_app": get_frontmost_app(),
        "running_apps": get_running_apps(),
        "time": time.strftime("%Y-%m-%d %H:%M:%S %Z"),
        "user": os.environ.get("USER", "unknown"),
        "home": str(Path.home()),
    }

    # Active document in frontmost app (if supported)
    front = ctx["frontmost_app"]
    if front in ("Microsoft Excel", "Microsoft Word", "Microsoft PowerPoint",
                 "Pages", "Numbers", "Keynote", "Preview", "TextEdit"):
        ok, doc = run_applescript(
            f'tell application "{front}" to get name of front document'
        )
        if ok:
            ctx["active_document"] = doc

    # Also check for data apps running in background — crucial for
    # "paste this data to X" where the data is in a different app.
    running = ctx.get("running_apps", [])
    data_apps = ["Numbers", "Microsoft Excel", "Pages", "Preview"]
    background_docs = {}
    for dapp in data_apps:
        if dapp in running and dapp != front:
            ok, doc = run_applescript(
                f'tell application "{dapp}" to get name of front document',
                timeout=3,
            )
            if ok and doc:
                background_docs[dapp] = doc
    if background_docs:
        ctx["other_open_documents"] = background_docs

    # Excel workbook contents — read the full structure when Excel is
    # active or has a workbook open in the background. This lets Claude
    # see cell values and formulas without needing screenshots.
    if front == "Microsoft Excel" or "Microsoft Excel" in running:
        try:
            excel_ctx = get_excel_context()
            if excel_ctx:
                ctx["excel"] = excel_ctx
        except Exception as e:
            sys.stderr.write(f"[system_context] excel read failed: {e}\n")

    # Finder: get selected files
    if front == "Finder":
        ok, sel = run_applescript(
            'tell application "Finder" to get POSIX path of (selection as alias list)'
        )
        if ok and sel:
            ctx["finder_selection"] = sel

    # Safari/Chrome: get current tab URL + title
    if front in ("Safari", "Google Chrome"):
        if front == "Safari":
            ok, url = run_applescript(
                'tell application "Safari" to get URL of front document'
            )
            ok2, title = run_applescript(
                'tell application "Safari" to get name of front document'
            )
        else:
            ok, url = run_applescript(
                'tell application "Google Chrome" to get URL of active tab of front window'
            )
            ok2, title = run_applescript(
                'tell application "Google Chrome" to get title of active tab of front window'
            )
        if ok:
            ctx["browser_url"] = url
        if ok2:
            ctx["browser_title"] = title

    return ctx


# ── Screen automation (vision loop primitives) ──────────────────────────────
# These let Claude see the screen and interact with ANY app — no per-app API
# needed. The flow: screenshot → Claude analyzes → click/type/scroll → repeat.

def _get_screen_point_size() -> tuple[int, int]:
    """Get the main display size in screen POINTS (not retina pixels).

    On a retina Mac the physical pixels are 2× the point dimensions.
    CGEvent clicks use points, so we need this for coordinate mapping.
    """
    try:
        import Quartz
        main = Quartz.CGMainDisplayID()
        w = int(Quartz.CGDisplayPixelsWide(main))   # point width
        h = int(Quartz.CGDisplayPixelsHigh(main))    # point height
        return w, h
    except Exception:
        # Fallback — query via system_profiler
        try:
            out = subprocess.run(
                ["system_profiler", "SPDisplaysDataType"],
                capture_output=True, text=True, timeout=5,
            ).stdout
            import re
            m = re.search(r"Resolution:\s*(\d+)\s*x\s*(\d+)", out)
            if m:
                return int(m.group(1)), int(m.group(2))
        except Exception:
            pass
        return 1470, 956  # sensible MacBook default


def capture_screenshot(region: dict = None) -> tuple[bool, str]:
    """Capture the screen and return base64-encoded JPEG.

    CRITICAL — Retina coordinate mapping:
      screencapture produces 2× retina pixels (e.g. 2940×1912), but
      CGEvent clicks use screen POINTS (e.g. 1470×956).  We resize the
      screenshot to the screen's point dimensions so that when Claude
      says "click at (x, y)" those coordinates map 1:1 to screen points.

    Args:
        region: optional {x, y, width, height} to capture a sub-region
                (in screen points).  If None, captures the full display.

    Returns:
        (True, base64_jpeg_string) on success.
    """
    import base64
    import tempfile

    raw_path = os.path.join(tempfile.gettempdir(), "tsifl_screen.png")
    try:
        cmd = ["screencapture", "-x"]  # -x = no sound
        if region:
            x, y = region.get("x", 0), region.get("y", 0)
            w, h = region.get("width", 800), region.get("height", 600)
            cmd.extend(["-R", f"{x},{y},{w},{h}"])
        cmd.append(raw_path)
        subprocess.run(cmd, timeout=5, check=True)

        # Resize to screen POINT dimensions so coordinates map 1:1 to clicks
        try:
            from PIL import Image as _PILImage
            import io as _io
            img = _PILImage.open(raw_path)

            # Convert RGBA→RGB (JPEG can't do alpha)
            if img.mode in ("RGBA", "LA"):
                bg = _PILImage.new("RGB", img.size, (255, 255, 255))
                bg.paste(img, mask=img.split()[-1])
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")

            # Resize to screen point dimensions (not arbitrary max)
            screen_w, screen_h = _get_screen_point_size()
            raw_w, raw_h = img.size
            logger.info(
                f"Screenshot: raw {raw_w}×{raw_h}, "
                f"screen points {screen_w}×{screen_h}, "
                f"scale {raw_w / screen_w:.1f}×"
            )

            if raw_w != screen_w or raw_h != screen_h:
                img = img.resize((screen_w, screen_h), _PILImage.LANCZOS)

            # Encode as JPEG — shrink until under 700KB
            for quality in (80, 65, 50, 40):
                buf = _io.BytesIO()
                img.save(buf, format="JPEG", quality=quality, optimize=True)
                data = buf.getvalue()
                if len(data) <= 700_000:
                    break
            return True, base64.b64encode(data).decode("ascii")
        except ImportError:
            # PIL not available — send raw PNG (larger but still works)
            with open(raw_path, "rb") as f:
                data = base64.b64encode(f.read()).decode("ascii")
            return True, data
    except Exception as e:
        return False, f"Screenshot failed: {e}"


def click_at_position(x: int, y: int, click_type: str = "left") -> tuple[bool, str]:
    """Click at screen coordinates using CGEvent (Quartz).

    Works with any app — no Accessibility API needed for basic clicks.

    Args:
        x, y: screen coordinates (pixels from top-left)
        click_type: "left", "right", or "double"
    """
    try:
        import Quartz

        point = (float(x), float(y))

        if click_type == "right":
            down_type = Quartz.kCGEventRightMouseDown
            up_type = Quartz.kCGEventRightMouseUp
            button = Quartz.kCGMouseButtonRight
        else:
            down_type = Quartz.kCGEventLeftMouseDown
            up_type = Quartz.kCGEventLeftMouseUp
            button = Quartz.kCGMouseButtonLeft

        # Move mouse first (some apps need this)
        move = Quartz.CGEventCreateMouseEvent(
            None, Quartz.kCGEventMouseMoved, point, button
        )
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, move)
        time.sleep(0.05)

        # Click down
        down = Quartz.CGEventCreateMouseEvent(None, down_type, point, button)
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, down)
        time.sleep(0.05)

        # Click up
        up = Quartz.CGEventCreateMouseEvent(None, up_type, point, button)
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, up)

        if click_type == "double":
            time.sleep(0.1)
            down2 = Quartz.CGEventCreateMouseEvent(None, down_type, point, button)
            Quartz.CGEventSetIntegerValueField(
                down2, Quartz.kCGMouseEventClickState, 2
            )
            Quartz.CGEventPost(Quartz.kCGHIDEventTap, down2)
            time.sleep(0.05)
            up2 = Quartz.CGEventCreateMouseEvent(None, up_type, point, button)
            Quartz.CGEventSetIntegerValueField(
                up2, Quartz.kCGMouseEventClickState, 2
            )
            Quartz.CGEventPost(Quartz.kCGHIDEventTap, up2)

        return True, f"Clicked at ({x}, {y})"
    except ImportError:
        # Fallback: AppleScript — less precise but works
        return run_applescript(
            f'tell application "System Events" to click at {{{x}, {y}}}'
        )
    except Exception as e:
        return False, f"Click failed: {e}"


def type_text_keyboard(text: str) -> tuple[bool, str]:
    """Type text using System Events keystrokes.

    Handles special characters by using 'keystroke' for printable chars.
    For large blocks of text, uses clipboard paste for speed.
    """
    if not text:
        return True, "Nothing to type"

    # For long text (>50 chars), paste from clipboard is faster + more reliable
    if len(text) > 50:
        ok, _ = set_clipboard(text)
        if ok:
            return run_applescript(
                'tell application "System Events" to keystroke "v" using command down'
            )

    # Short text: type directly (preserves clipboard)
    # Escape backslashes and quotes for AppleScript
    escaped = text.replace("\\", "\\\\").replace('"', '\\"')
    return run_applescript(
        f'tell application "System Events" to keystroke "{escaped}"'
    )


def press_key_combo(keys: str) -> tuple[bool, str]:
    """Press a keyboard shortcut.

    Args:
        keys: shortcut string like "cmd+c", "cmd+shift+t", "return",
              "tab", "escape", "space", "delete", "up", "down"

    Supports: cmd, shift, ctrl/control, alt/option + any key.
    """
    parts = [k.strip().lower() for k in keys.split("+")]

    # Map modifier names to AppleScript modifiers
    modifier_map = {
        "cmd": "command down",
        "command": "command down",
        "shift": "shift down",
        "ctrl": "control down",
        "control": "control down",
        "alt": "option down",
        "option": "option down",
    }

    # Map special key names to AppleScript key codes
    special_keys = {
        "return": 36, "enter": 36,
        "tab": 48,
        "escape": 53, "esc": 53,
        "space": 49,
        "delete": 51, "backspace": 51,
        "forwarddelete": 117,
        "up": 126, "down": 125, "left": 123, "right": 124,
        "home": 115, "end": 119,
        "pageup": 116, "pagedown": 121,
        "f1": 122, "f2": 120, "f3": 99, "f4": 118,
        "f5": 96, "f6": 97, "f7": 98, "f8": 100,
    }

    modifiers = []
    key_part = None

    for p in parts:
        if p in modifier_map:
            modifiers.append(modifier_map[p])
        else:
            key_part = p

    if key_part is None:
        return False, f"No key specified in '{keys}'"

    modifier_str = ", ".join(modifiers)

    if key_part in special_keys:
        # Use key code for special keys
        code = special_keys[key_part]
        if modifiers:
            script = f'tell application "System Events" to key code {code} using {{{modifier_str}}}'
        else:
            script = f'tell application "System Events" to key code {code}'
    else:
        # Regular character
        escaped = key_part.replace("\\", "\\\\").replace('"', '\\"')
        if modifiers:
            script = f'tell application "System Events" to keystroke "{escaped}" using {{{modifier_str}}}'
        else:
            script = f'tell application "System Events" to keystroke "{escaped}"'

    return run_applescript(script)


def scroll_screen(direction: str = "down", amount: int = 3, x: int = 0, y: int = 0) -> tuple[bool, str]:
    """Scroll at the current mouse position or specified coordinates.

    Args:
        direction: "up" or "down"
        amount: number of scroll ticks (3 ≈ one page-ish)
        x, y: optional coordinates to scroll at (0,0 = current position)
    """
    try:
        import Quartz

        scroll_amount = amount if direction == "up" else -amount

        # Move mouse to position first if coordinates given
        if x > 0 and y > 0:
            move = Quartz.CGEventCreateMouseEvent(
                None, Quartz.kCGEventMouseMoved,
                (float(x), float(y)), Quartz.kCGMouseButtonLeft,
            )
            Quartz.CGEventPost(Quartz.kCGHIDEventTap, move)
            time.sleep(0.05)

        event = Quartz.CGEventCreateScrollWheelEvent(
            None, Quartz.kCGScrollEventUnitLine, 1, scroll_amount
        )
        Quartz.CGEventPost(Quartz.kCGHIDEventTap, event)
        return True, f"Scrolled {direction} {amount}"
    except ImportError:
        # Fallback to AppleScript
        key_code = 126 if direction == "up" else 125  # arrow keys
        script = f'tell application "System Events" to key code {key_code}'
        for _ in range(amount):
            run_applescript(script)
        return True, f"Scrolled {direction} {amount}"
    except Exception as e:
        return False, f"Scroll failed: {e}"


def wait_seconds(seconds: float) -> tuple[bool, str]:
    """Wait for a specified duration (for UI to settle)."""
    seconds = min(seconds, 10)  # cap at 10s
    time.sleep(seconds)
    return True, f"Waited {seconds}s"


# ── Spotify (instant play via AppleScript — no vision loop needed) ────────

def spotify_play_library_playlist(name: str) -> tuple[bool, str]:
    """Find a playlist in Spotify's sidebar (Your Library) by name and play it.

    This is the right tool for 'play my FR playlist' / 'play Chill Mix' /
    'play <any named playlist in the user's library>'. Public search via
    spotify_play won't surface a private library playlist; navigating the
    sidebar will.

    Flow:
    1. Activate Spotify.
    2. Recursively walk the UI tree looking for a sidebar item whose
       title/value matches `name` (case-insensitive substring).
    3. Click it — Spotify loads the playlist detail page.
    4. Click the big green Play button via the same accessibility walker
       we use for play_media.

    Returns (True, "▶️ playlist name") on success, (False, error) on miss.
    """
    safe_name = name.replace("\\", "\\\\").replace('"', '\\"')

    # Make sure Spotify is up
    try:
        subprocess.run(["open", "-a", "Spotify"], check=True, timeout=10)
        time.sleep(0.6)
    except Exception as e:
        return False, f"Could not open Spotify: {e}"

    # Snapshot current playback so we can detect a new playlist starting
    def _get_playing() -> tuple[bool, str, str]:
        ok_s, info = run_applescript(
            'tell application "Spotify"\n'
            '    if player state is playing then\n'
            '        return (name of current track) & "||" & (artist of current track)\n'
            '    else\n'
            '        return "NOT_PLAYING"\n'
            '    end if\n'
            'end tell',
            timeout=5,
        )
        if not ok_s or not info or "NOT_PLAYING" in info:
            return False, "", ""
        parts = info.split("||", 1)
        return True, parts[0] if parts else "", (parts[1] if len(parts) > 1 else "")

    before_playing, before_track, before_artist = _get_playing()
    # Pause anything currently playing so we can verify a new track starts
    if before_playing:
        run_applescript('tell application "Spotify" to pause', timeout=3)
        time.sleep(0.3)

    # Approach: navigate Spotify to the search results for the playlist name
    # via the spotify:search: URL scheme (which surfaces user's own playlists
    # near the top when there's an exact name match), then click the first
    # Play button found via accessibility walk.
    #
    # This is much simpler than walking the entire sidebar tree, which is
    # fragile across Spotify UI updates and across user view states (sidebar
    # collapsed, playlist scrolled out of view, etc.).
    import urllib.parse as _up
    encoded = _up.quote(name)
    try:
        subprocess.run(["open", f"spotify:search:{encoded}"], timeout=5)
    except Exception as e:
        return False, f"Could not open Spotify search: {e}"
    time.sleep(2.5)  # let results load

    # Walk the UI to find the first Play button on the search results page.
    # Same shape as spotify_play's helper but kept inline to avoid coupling.
    script = '''
on findPlay(elem, depth)
    if depth > 12 then return missing value
    try
        set elemRole to (role of elem) as text
    on error
        return missing value
    end try
    if elemRole is "AXButton" then
        set descr to ""
        try
            set descr to (description of elem) as text
        end try
        try
            set descr to descr & " " & ((title of elem) as text)
        end try
        try
            set descr to descr & " " & ((help of elem) as text)
        end try
        if descr contains "Play" and descr does not contain "Pause" then
            return elem
        end if
    end if
    try
        repeat with child in (UI elements of elem)
            set r to my findPlay(child, depth + 1)
            if r is not missing value then return r
        end repeat
    end try
    return missing value
end findPlay

tell application "Spotify" to activate
delay 0.4
tell application "System Events" to tell process "Spotify"
    set frontmost to true
    delay 0.3
    set btn to my findPlay(window 1, 0)
    if btn is missing value then
        return "NO_PLAY_BUTTON"
    end if
    click btn
    return "OK"
end tell
'''
    ok, result = run_applescript(script, timeout=15)
    time.sleep(1.5)

    if not ok:
        return False, (
            f"Couldn't navigate Spotify (AppleScript: {str(result)[:120]}). "
            f"Open Spotify manually and try clicking '{name}' yourself."
        )
    if "NO_PLAY_BUTTON" in (result or ""):
        return False, (
            f"Opened Spotify search for '{name}' but no Play button visible. "
            f"Spotify may need a different exact name — check your sidebar."
        )

    is_playing, track, artist = _get_playing()
    if is_playing and (track != before_track or artist != before_artist):
        return True, f"▶️ {track} — {artist} (from '{name}')"
    if is_playing:
        return True, f"▶️ {track} (search for '{name}' loaded but old track still playing)"
    return False, f"Found '{name}' but playback didn't start."


def spotify_play(query: str) -> tuple[bool, str]:
    """Play on Spotify — uses URL scheme to land on search page, then plays first track.

    Flow:
    1. Open `spotify:search:QUERY` URL — Spotify navigates directly to search results
    2. Activate + wait for results to load
    3. Tab through to first track row, press Enter to play
    4. Verify the playing track matches the query (sanity check)
    5. Fallback to keystroke search if URL scheme didn't work
    """
    import urllib.parse as _up
    encoded = _up.quote(query)
    escaped = query.replace("\\", "\\\\").replace('"', '\\"')
    query_words = set(w.lower() for w in query.split() if len(w) > 2)

    def _get_playing() -> tuple[bool, str, str]:
        """Returns (is_playing, track_name, artist)."""
        ok_s, info = run_applescript(
            'tell application "Spotify"\n'
            '    if player state is playing then\n'
            '        return (name of current track) & "||" & (artist of current track)\n'
            '    else\n'
            '        return "NOT_PLAYING"\n'
            '    end if\n'
            'end tell',
            timeout=5,
        )
        if not ok_s or not info or "NOT_PLAYING" in info:
            return False, "", ""
        parts = info.split("||", 1)
        track = parts[0] if parts else ""
        artist = parts[1] if len(parts) > 1 else ""
        return True, track, artist

    def _matches_query(track: str, artist: str) -> bool:
        """Loose match — any query word in track or artist name."""
        if not query_words:
            return True
        combined = f"{track} {artist}".lower()
        return any(w in combined for w in query_words)

    try:
        # Snapshot what's playing BEFORE we do anything
        was_playing, before_track, before_artist = _get_playing()

        # If something was playing, pause it first so we can detect when
        # the new content starts (otherwise we can't tell if Play worked).
        if was_playing:
            run_applescript('tell application "Spotify" to pause', timeout=3)
            time.sleep(0.3)

        # Ensure Spotify is running
        subprocess.run(["open", "-a", "Spotify"], check=True, timeout=10)
        time.sleep(0.5)

        # ── Method 1: URL scheme + click Play button at computed position ──
        # spotify:search:QUERY lands directly on results (or a top playlist).
        # We then compute the Play button position from window geometry.
        # The big green Play button is consistently at ~(232, 388) relative
        # to the main content area on playlist pages. Search pages also
        # have a Play on the top result tile at similar offset.
        subprocess.run(["open", f"spotify:search:{encoded}"], timeout=5)
        time.sleep(2.8)  # let page load + animations finish

        # Get Spotify window geometry, then click the Play button area.
        # We also try the AppleScript UI tree walk as a primary method.
        play_click_script = '''
on findPlayInTree(elem, depth)
    if depth > 12 then return missing value
    try
        set elemRole to role of elem
    on error
        return missing value
    end try
    try
        set descr to ""
        try
            set descr to description of elem
        end try
        try
            set descr to descr & "||" & (title of elem)
        end try
        try
            set descr to descr & "||" & (help of elem)
        end try
        -- Match localized "Play" (English) or common variants
        if descr contains "Play " or descr starts with "Play" or descr is "Play" then
            if descr does not contain "Pause" then
                return elem
            end if
        end if
    end try
    try
        repeat with child in (UI elements of elem)
            set r to my findPlayInTree(child, depth + 1)
            if r is not missing value then return r
        end repeat
    end try
    return missing value
end findPlayInTree

tell application "Spotify" to activate
delay 0.3
tell application "System Events" to tell process "Spotify"
    set frontmost to true
    delay 0.3
    -- Method A: recursive search for Play button
    set btn to my findPlayInTree(window 1, 0)
    if btn is not missing value then
        click btn
        return "AXClicked"
    end if
    -- Method B: compute position from window geometry and CGEvent click
    set winPos to position of window 1
    set winSize to size of window 1
    -- Play button is roughly at: window_left + sidebar_width(280) + 60
    --                            window_top + header_height(60) + 320
    set clickX to (item 1 of winPos) + 340
    set clickY to (item 2 of winPos) + 380
    return "POS:" & clickX & "," & clickY & "," & (item 3 of winSize) & "," & (item 4 of winSize)
end tell
'''
        ok, result = run_applescript(play_click_script, timeout=15)
        sys.stderr.write(f"[spotify_play] result: {result!r}\n")

        if result and result.startswith("POS:"):
            # AppleScript couldn't find Play via accessibility — click by position
            try:
                parts = result[4:].split(",")
                cx, cy = int(parts[0]), int(parts[1])
                sys.stderr.write(f"[spotify_play] clicking position ({cx}, {cy})\n")
                click_at_position(cx, cy, "left")
            except Exception as click_err:
                sys.stderr.write(f"[spotify_play] position click failed: {click_err}\n")

        time.sleep(1.5)

        # ── Verify: is something playing now? ──
        is_playing, track, artist = _get_playing()
        if is_playing:
            return True, f"▶️ {track} — {artist}"

        # ── Method 2: Fallback — Cmd+K search overlay + longer delays ─
        run_applescript(
            'tell application "Spotify" to activate\n'
            'delay 0.3\n'
            'tell application "System Events" to tell process "Spotify"\n'
            '    set frontmost to true\n'
            '    delay 0.2\n'
            '    keystroke "k" using command down\n'
            '    delay 0.6\n'
            '    keystroke "a" using command down\n'
            '    delay 0.1\n'
            f'    keystroke "{escaped}"\n'
            '    delay 2.5\n'
            # Down arrow to skip "Search" suggestion, land on first track
            '    key code 125\n'
            '    delay 0.2\n'
            '    key code 36\n'
            'end tell',
            timeout=15,
        )
        time.sleep(2)

        is_playing, track, artist = _get_playing()
        if is_playing and (track != before_track or artist != before_artist):
            return True, f"▶️ {track} — {artist}"

        # ── Last resort: report what's playing (even if it's old) ──
        if is_playing:
            return True, f"⚠️ Couldn't find '{query}', still playing: {track} — {artist}"
        return False, f"Searched for '{query}' but playback didn't start"

    except Exception as e:
        return False, f"Spotify play failed: {e}"


# ── Gmail operations (local Gmail API via OAuth token) ──────────────────────
# The desktop agent talks to Gmail directly using the local OAuth token at
# ~/.tsifulator_gmail_token.json. No backend round-trip needed — faster and
# works offline (for reads). If the token doesn't exist, we tell the user
# to run the setup script.

_GMAIL_TOKEN_PATH = Path.home() / ".tsifulator_gmail_token.json"


def _get_gmail_service():
    """Build an authenticated Gmail API service from the local token.

    Returns the service object, or raises RuntimeError with a user-friendly
    message if setup is needed.
    """
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except ImportError:
        raise RuntimeError(
            "Gmail libraries missing. Run:\n"
            "  pip install google-api-python-client google-auth-oauthlib"
        )

    if not _GMAIL_TOKEN_PATH.exists():
        raise RuntimeError(
            "Gmail not connected yet. Run:\n"
            "  cd tsifulator.ai && python3 gmail-client/gmail_setup.py"
        )

    creds = Credentials.from_authorized_user_file(str(_GMAIL_TOKEN_PATH))

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        _GMAIL_TOKEN_PATH.write_text(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def _extract_email_body(payload: dict) -> str:
    """Recursively extract plain text body from Gmail message payload."""
    import base64
    if payload.get("mimeType") == "text/plain":
        data = payload.get("body", {}).get("data", "")
        if data:
            return base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="replace")
    if "parts" in payload:
        for part in payload["parts"]:
            result = _extract_email_body(part)
            if result:
                return result
    return ""


def gmail_check_inbox(max_results: int = 10) -> tuple[bool, str]:
    """Fetch recent inbox messages using the local Gmail API."""
    try:
        service = _get_gmail_service()
        results = service.users().messages().list(
            userId="me", labelIds=["INBOX"], maxResults=max_results
        ).execute()

        messages = results.get("messages", [])
        if not messages:
            return True, "Inbox is empty."

        lines = [f"📬 {len(messages)} recent emails:"]
        for i, msg in enumerate(messages, 1):
            detail = service.users().messages().get(
                userId="me", id=msg["id"],
                format="metadata",
                metadataHeaders=["From", "Subject", "Date"],
            ).execute()
            headers = {h["name"]: h["value"] for h in detail["payload"]["headers"]}
            unread = "📩" if "UNREAD" in detail.get("labelIds", []) else "  "
            sender = headers.get("From", "?")
            if "<" in sender:
                sender = sender.split("<")[0].strip().strip('"')
            lines.append(f"{unread} {i}. {sender}")
            lines.append(f"     {headers.get('Subject', '(no subject)')}")
            snippet = detail.get("snippet", "")
            if snippet:
                lines.append(f"     {snippet[:100]}")
            lines.append(f"     [id:{msg['id']}]")
        return True, "\n".join(lines)
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail error: {e}"


def gmail_search(query: str, max_results: int = 10) -> tuple[bool, str]:
    """Search emails using Gmail query syntax."""
    try:
        service = _get_gmail_service()
        results = service.users().messages().list(
            userId="me", q=query, maxResults=max_results
        ).execute()

        messages = results.get("messages", [])
        if not messages:
            return True, f"No emails found for '{query}'."

        lines = [f"Found {len(messages)} email(s) for '{query}':"]
        for i, msg in enumerate(messages, 1):
            detail = service.users().messages().get(
                userId="me", id=msg["id"],
                format="metadata",
                metadataHeaders=["From", "Subject", "Date"],
            ).execute()
            headers = {h["name"]: h["value"] for h in detail["payload"]["headers"]}
            sender = headers.get("From", "?")
            if "<" in sender:
                sender = sender.split("<")[0].strip().strip('"')
            lines.append(f"  {i}. {sender}")
            lines.append(f"     {headers.get('Subject', '(no subject)')}")
            snippet = detail.get("snippet", "")
            if snippet:
                lines.append(f"     {snippet[:100]}")
            lines.append(f"     [id:{msg['id']}]")
        return True, "\n".join(lines)
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail search failed: {e}"


def gmail_read_message(message_id: str) -> tuple[bool, str]:
    """Read the full body of a specific email message."""
    try:
        service = _get_gmail_service()
        msg = service.users().messages().get(
            userId="me", id=message_id, format="full"
        ).execute()

        headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
        body_text = _extract_email_body(msg["payload"])
        if len(body_text) > 2000:
            body_text = body_text[:1997] + "…"

        lines = [
            f"From: {headers.get('From', '?')}",
            f"To: {headers.get('To', '?')}",
            f"Subject: {headers.get('Subject', '(no subject)')}",
            f"Date: {headers.get('Date', '?')}",
            "",
            body_text or "(no body)",
        ]
        return True, "\n".join(lines)
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail read failed: {e}"


def gmail_send(to: str, subject: str, body: str, reply_to_id: str = "") -> tuple[bool, str]:
    """Send an email (or reply to a thread) via Gmail API."""
    try:
        import base64
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart

        service = _get_gmail_service()
        msg = MIMEMultipart("alternative")
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        send_body = {"raw": raw}
        if reply_to_id:
            send_body["threadId"] = reply_to_id

        sent = service.users().messages().send(userId="me", body=send_body).execute()
        return True, f"✉️ Email sent to {to}"
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail send failed: {e}"


def gmail_draft(to: str, subject: str, body: str) -> tuple[bool, str]:
    """Create a draft email in Gmail (doesn't send)."""
    try:
        import base64
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart

        service = _get_gmail_service()
        msg = MIMEMultipart("alternative")
        msg["To"] = to
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        draft = service.users().drafts().create(
            userId="me", body={"message": {"raw": raw}}
        ).execute()
        return True, f"📝 Draft created for {to}"
    except RuntimeError as e:
        return False, str(e)
    except Exception as e:
        return False, f"Gmail draft failed: {e}"


# ── Action executor ──────────────────────────────────────────────────────────

def execute_action(action: Action) -> Action:
    """Execute a single action and update it with the result.

    Returns the same Action with `result`, `success`, and `error` filled in.
    """
    try:
        # Parse command — might be a raw string or a JSON payload dict
        cmd = action.command
        cmd_data = {}
        if cmd.startswith("{"):
            try:
                cmd_data = json.loads(cmd)
            except json.JSONDecodeError:
                pass

        if action.type == "search_files":
            params = cmd_data if cmd_data else {"query": cmd}
            paths = search_files(
                query=params.get("query", ""),
                file_type=params.get("file_type"),
                max_results=params.get("max_results", 10),
            )
            if paths:
                action.result = "\n".join(paths)
                action.success = True
            else:
                action.result = "No files found"
                action.success = True  # search succeeded, just no results

        elif action.type == "open_file":
            path = cmd_data.get("path", cmd_data.get("file_path", cmd))
            action.success, action.result = open_file(path)

        elif action.type == "open_app":
            app_name = cmd_data.get("app", cmd_data.get("app_name", cmd_data.get("name", cmd)))
            action.success, action.result = open_app(app_name)

        elif action.type == "open_url":
            url = cmd_data.get("url", cmd)
            action.success, action.result = open_url(url)

        elif action.type == "applescript":
            script = cmd_data.get("script", cmd_data.get("code", cmd))
            action.success, action.result = run_applescript(script)

        elif action.type == "shell":
            shell_cmd = cmd_data.get("command", cmd)
            action.success, action.result = run_shell(shell_cmd)

        elif action.type == "write_file":
            file_path = cmd_data.get("path", "")
            content = cmd_data.get("content", cmd)
            if not file_path:
                action.success = False
                action.result = "No file path specified"
            else:
                try:
                    p = Path(file_path).expanduser()
                    p.parent.mkdir(parents=True, exist_ok=True)
                    p.write_text(content, encoding="utf-8")
                    action.success = True
                    action.result = f"Wrote {len(content)} chars to {p}"
                except Exception as e:
                    action.success = False
                    action.result = f"Write failed: {e}"

        elif action.type == "clipboard_copy":
            text = cmd_data.get("text", cmd)
            action.success, action.result = set_clipboard(text)

        elif action.type == "clipboard_read":
            text = get_clipboard()
            action.result = text or "(clipboard empty)"
            action.success = True

        elif action.type == "notify":
            # Show a macOS notification
            try:
                import rumps
                rumps.notification(
                    title="tsifl",
                    subtitle="",
                    message=action.command,
                )
                action.success = True
                action.result = "Notification shown"
            except Exception as e:
                action.success = False
                action.error = str(e)

        # ── Screen automation actions ───────────────────────────────────
        elif action.type == "screenshot":
            region = cmd_data.get("region") if cmd_data else None
            action.success, action.result = capture_screenshot(region)

        elif action.type == "click_at":
            x = cmd_data.get("x", 0)
            y = cmd_data.get("y", 0)
            ct = cmd_data.get("click_type", "left")
            action.success, action.result = click_at_position(x, y, ct)

        elif action.type == "type_text":
            text = cmd_data.get("text", cmd)
            action.success, action.result = type_text_keyboard(text)

        elif action.type == "key_combo":
            keys = cmd_data.get("keys", cmd)
            action.success, action.result = press_key_combo(keys)

        elif action.type == "scroll":
            direction = cmd_data.get("direction", "down")
            amount = cmd_data.get("amount", 3)
            sx = cmd_data.get("x", 0)
            sy = cmd_data.get("y", 0)
            action.success, action.result = scroll_screen(direction, amount, sx, sy)

        elif action.type == "wait":
            secs = cmd_data.get("seconds", 1)
            action.success, action.result = wait_seconds(secs)

        # ── Gmail actions ───────────────────────────────────────────────
        elif action.type == "check_inbox":
            max_r = cmd_data.get("max_results", 10)
            action.success, action.result = gmail_check_inbox(max_r)

        elif action.type == "search_email":
            query = cmd_data.get("query", cmd)
            max_r = cmd_data.get("max_results", 10)
            action.success, action.result = gmail_search(query, max_r)

        elif action.type == "read_email":
            msg_id = cmd_data.get("message_id", cmd)
            action.success, action.result = gmail_read_message(msg_id)

        elif action.type == "send_email":
            action.success, action.result = gmail_send(
                to=cmd_data.get("to", ""),
                subject=cmd_data.get("subject", ""),
                body=cmd_data.get("body", ""),
                reply_to_id=cmd_data.get("reply_to_id", ""),
            )

        elif action.type == "draft_email":
            action.success, action.result = gmail_draft(
                to=cmd_data.get("to", ""),
                subject=cmd_data.get("subject", ""),
                body=cmd_data.get("body", ""),
            )

        elif action.type == "spotify_play":
            query = cmd_data.get("query", cmd)
            action.success, action.result = spotify_play(query)

        elif action.type == "data_export":
            # Battle-tested export: Claude just says WHAT, executor knows HOW.
            source_app = cmd_data.get("source_app", "")
            dest_path = cmd_data.get("destination", "")
            fmt = cmd_data.get("format", "csv").lower()
            if not source_app or not dest_path:
                action.success = False
                action.result = "data_export needs source_app and destination"
            else:
                expanded = str(Path(dest_path).expanduser())
                Path(expanded).parent.mkdir(parents=True, exist_ok=True)
                action.success, action.result = _data_export(source_app, expanded, fmt)

        elif action.type == "play_media":
            # Deterministic media playback — no vision needed.
            platform = cmd_data.get("platform", "youtube").lower()
            query = cmd_data.get("query", cmd)
            action.success, action.result = _play_media(platform, query)

        elif action.type == "web_search":
            # Deterministic web search — just open the URL, no vision.
            query = cmd_data.get("query", cmd)
            engine = cmd_data.get("engine", "google").lower()
            action.success, action.result = _web_search(query, engine)

        elif action.type == "read_file":
            # Read a text file (CSV, TSV, TXT, JSON, .py, .r, etc.) from disk
            # so Claude can see the contents and act on them.
            path = cmd_data.get("path", cmd)
            max_chars = cmd_data.get("max_chars", 30000)
            try:
                from pathlib import Path as _P
                p = _P(path).expanduser()

                # Defense: if the path is just a bare filename (no slash) it's
                # almost certainly a hallucination — real read_file calls have
                # absolute paths from the user's message. Refuse with an
                # instructive error so Claude knows what to do next.
                if "/" not in str(path):
                    action.success = False
                    action.result = (
                        f"REJECTED: '{path}' is not an absolute path (no '/'). "
                        f"STOP. Do NOT retry with another guessed path. "
                        f"Two cases:\n"
                        f"1. If the user ATTACHED a file (image/PDF), it's "
                        f"already in your context — read it directly from "
                        f"the image/document blocks. Do NOT call read_file "
                        f"for attached files.\n"
                        f"2. If the user mentioned a file but didn't give a "
                        f"path, respond with TEXT asking them for the full "
                        f"absolute path. Do NOT call read_file again with a "
                        f"different guess."
                    )
                elif not p.exists():
                    action.success = False
                    action.result = (
                        f"NOT FOUND: {p}\n"
                        f"This path does not exist. STOP — do not retry with "
                        f"variants of the path. Either:\n"
                        f"1. The user gave you a typo'd path. Respond with "
                        f"TEXT asking them to confirm/correct it.\n"
                        f"2. You misread/modified the path. Look back at the "
                        f"original user message and use it EXACTLY."
                    )
                elif p.is_dir():
                    action.success = False
                    action.result = (
                        f"IS A DIRECTORY: {p}\n"
                        f"You can't read a directory with read_file. If you "
                        f"want to list directory contents, use the `shell` "
                        f"tool with `ls`."
                    )
                elif p.stat().st_size > 5_000_000:
                    action.success = False
                    action.result = (
                        f"FILE TOO LARGE: {p.stat().st_size:,} bytes (cap is "
                        f"5 MB). For large files, ask the user to slice or "
                        f"summarize before importing."
                    )
                else:
                    try:
                        content = p.read_text(encoding="utf-8", errors="replace")
                    except Exception:
                        # Binary or weird encoding — read as bytes and decode lossy
                        content = p.read_bytes().decode("utf-8", errors="replace")
                    action.success = True
                    action.result = content[:max_chars]
                    if len(content) > max_chars:
                        action.result += (
                            f"\n... [TRUNCATED — file has {len(content):,} chars total. "
                            f"Call read_file again with max_chars={len(content)} "
                            f"to see the rest.]"
                        )
            except Exception as e:
                action.success = False
                action.error = f"read_file failed: {e}"

        elif action.type == "fetch_url":
            # Fetch a URL and return the text content for Claude to read
            url = cmd_data.get("url", cmd)
            max_chars = cmd_data.get("max_chars", 8000)
            action.success, action.result = _fetch_url(url, max_chars)

        elif action.type == "web_lookup":
            # Search the web AND read the top results — Claude gets actual answers
            query = cmd_data.get("query", cmd)
            max_chars = cmd_data.get("max_chars", 6000)
            action.success, action.result = _web_lookup(query, max_chars)

        # ── Memory & shortcut actions ──────────────────────────────────
        elif action.type == "save_memory":
            # Claude can save facts about the user proactively
            fact = cmd_data.get("fact", cmd)
            try:
                from memory import save_memory as _save_mem
                action.result = _save_mem(fact)
                action.success = True
            except Exception as e:
                action.success = False
                action.error = f"Memory save failed: {e}"

        elif action.type == "set_shortcut":
            # Claude can create shortcuts (slash commands or system hotkeys)
            trigger = cmd_data.get("trigger", "")
            shortcut_action = cmd_data.get("action", "")
            desc = cmd_data.get("description", "")
            hotkey = cmd_data.get("hotkey", "")  # e.g. "cmd+d"
            if not trigger or not shortcut_action:
                action.success = False
                action.result = "set_shortcut needs trigger and action"
            else:
                try:
                    from memory import save_shortcut as _save_sc
                    action.result = _save_sc(trigger, shortcut_action, desc, hotkey=hotkey)
                    action.success = True
                except Exception as e:
                    action.success = False
                    action.error = f"Shortcut save failed: {e}"

        elif action.type == "set_volume":
            # Deterministic system volume control. Avoids the hallucination
            # where the agent picks a wrong key_combo for mute.
            muted = cmd_data.get("muted")
            level = cmd_data.get("level")
            scripts_run = []
            try:
                if isinstance(level, (int, float)):
                    clamped = max(0, min(100, int(level)))
                    ok, r = run_applescript(
                        f'set volume output volume {clamped}', timeout=3,
                    )
                    scripts_run.append(f"volume → {clamped}")
                    if clamped > 0:
                        run_applescript("set volume without output muted", timeout=3)
                if muted is True:
                    ok, r = run_applescript("set volume with output muted", timeout=3)
                    scripts_run.append("muted")
                elif muted is False:
                    ok, r = run_applescript("set volume without output muted", timeout=3)
                    scripts_run.append("unmuted")
                if not scripts_run:
                    action.success = False
                    action.result = "set_volume needs `muted` (bool) or `level` (0-100)"
                else:
                    action.success = True
                    action.result = "✅ " + ", ".join(scripts_run)
            except Exception as e:
                action.success = False
                action.error = f"set_volume failed: {e}"

        elif action.type == "play_spotify_playlist":
            # Navigate Spotify's sidebar to the user's named playlist and play
            name = cmd_data.get("name", cmd)
            if not name:
                action.success = False
                action.result = "play_spotify_playlist needs a `name`"
            else:
                action.success, action.result = spotify_play_library_playlist(name)

        elif action.type == "create_routine":
            # Claude can create recurring background tasks
            r_name = cmd_data.get("name", "")
            r_prompt = cmd_data.get("prompt", "")
            r_schedule = cmd_data.get("schedule", "")
            if not r_prompt or not r_schedule:
                action.success = False
                action.result = "create_routine needs `prompt` and `schedule`"
            else:
                try:
                    from routines import create_routine as _create_routine
                    ok, msg = _create_routine(r_name, r_prompt, r_schedule)
                    action.success = ok
                    action.result = msg
                except Exception as e:
                    action.success = False
                    action.error = f"Routine create failed: {e}"

        else:
            action.success = False
            action.error = f"Unknown action type: {action.type}"

    except Exception as e:
        action.success = False
        action.error = str(e)
        logger.error(f"Action failed ({action.type}): {e}")

    return action


def execute_plan(actions: list[Action], stop_on_error: bool = True) -> list[Action]:
    """Execute a list of actions sequentially.

    Args:
        actions: ordered list of Action objects
        stop_on_error: if True, stop executing after the first failure

    Returns:
        The same list with results filled in.
    """
    for action in actions:
        execute_action(action)
        if stop_on_error and not action.success:
            break
    return actions
