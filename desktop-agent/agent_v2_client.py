"""
agent_v2_client.py — talks to the v2 backend (/agent/turn + /agent/result).

The v2 flow:
  1. POST /agent/turn with the user's message + context + optional images
     → backend returns tool_uses (list of {id, name, input}) and text
  2. For each tool_use, run executor.execute_action() and collect results
  3. POST /agent/result with the tool_use_ids + their string outputs
     → backend returns more tool_uses, or done=true
  4. Loop until done=true

Conversation state lives on the backend, keyed by conversation_id.
This client just tracks the id between turns.
"""

import json
import os
import sys
import time
from pathlib import Path
from typing import Optional

import httpx

# Import bookkeeping from the main helper module — done lazily inside
# functions to avoid circular import on module load.

# ── Risk mapping per tool name (matches old behavior) ──────────────────────
_TOOL_RISK = {
    # GREEN — read-only
    "search_files": "green", "read_file": "green", "open_file": "green",
    "open_app": "green", "open_url": "green", "shell": "green",
    "clipboard_read": "green", "clipboard_copy": "green", "notify": "green",
    "play_media": "green", "web_open": "green", "fetch_url": "green",
    "check_inbox": "green", "search_email": "green", "read_email": "green",
    "screenshot": "green", "scroll": "green", "wait": "green",
    "save_memory": "green", "set_shortcut": "green",
    # YELLOW — writes/creates
    "applescript": "yellow", "write_file": "yellow", "data_export": "yellow",
    "click_at": "yellow", "type_text": "yellow", "key_combo": "yellow",
    "draft_email": "yellow",
    # RED — irreversible
    "send_email": "red",
}


def _backend_url() -> str:
    return os.getenv(
        "BACKEND_URL",
        "https://focused-solace-production-6839.up.railway.app",
    ).rstrip("/")


def _post_json(path: str, body: dict, timeout: int = 120) -> dict:
    """POST JSON to the backend with the same DNS fallback as the legacy client.

    Returns the parsed JSON response on success, or {"error": str} on failure.
    """
    url = f"{_backend_url()}{path}"

    # Reuse the DNS fallback logic from the helper module
    try:
        with httpx.Client(timeout=timeout) as client:
            r = client.post(url, json=body)
        return _parse_response(r)
    except Exception as dns_err:
        err_str = str(dns_err).lower()
        if "nodename" not in err_str and "errno 8" not in err_str:
            return {"error": str(dns_err)}

        # DNS failed — use cached IP fallback (same as legacy)
        from urllib.parse import urlparse
        hostname = urlparse(_backend_url()).hostname or ""
        ip = _get_cached_ip(hostname) or "66.33.22.247"

        import socket as _sock
        orig = _sock.getaddrinfo

        def patched(host, port, *args, **kwargs):
            if host == hostname:
                sys.stderr.write(f"[agent_v2] DNS override: {hostname} → {ip}\n")
                return orig(ip, port, *args, **kwargs)
            return orig(host, port, *args, **kwargs)

        _sock.getaddrinfo = patched
        try:
            with httpx.Client(timeout=timeout) as client:
                r = client.post(url, json=body)
            _save_cached_ip(hostname, ip)
            return _parse_response(r)
        except Exception as e:
            return {"error": str(e)}
        finally:
            _sock.getaddrinfo = orig


def _parse_response(r) -> dict:
    if r.status_code == 200:
        return r.json()
    return {"error": f"backend {r.status_code}: {r.text[:300]}"}


def _get_cached_ip(hostname: str) -> Optional[str]:
    try:
        cache_file = Path.home() / ".tsifl" / "dns_cache.json"
        if cache_file.exists():
            data = json.loads(cache_file.read_text())
            return data.get(hostname)
    except Exception:
        pass
    return None


def _save_cached_ip(hostname: str, ip: str):
    try:
        cache_file = Path.home() / ".tsifl" / "dns_cache.json"
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        data = {}
        if cache_file.exists():
            try:
                data = json.loads(cache_file.read_text())
            except Exception:
                data = {}
        data[hostname] = ip
        cache_file.write_text(json.dumps(data, indent=2))
    except Exception:
        pass


# ── Public API ──────────────────────────────────────────────────────────────

def start_turn(
    message: str,
    context: dict,
    conversation_id: Optional[str] = None,
    images: Optional[list[dict]] = None,
    model: Optional[str] = None,
) -> dict:
    """Send a new user message. Returns the turn response from /agent/turn.

    Response shape:
      {
        "conversation_id": str,
        "tool_uses": [{"id": str, "name": str, "input": dict}],
        "text": str,
        "thinking": str,
        "stop_reason": str,
        "done": bool,
        "usage": {...},
      }

    On error, returns {"error": str, "conversation_id": conversation_id, "tool_uses": [], "done": True}.
    """
    body = {
        "user_id": "shortcut-anon",
        "conversation_id": conversation_id,
        "message": message,
        "context": context,
        "images": images or [],
    }
    if model:
        body["model"] = model
    resp = _post_json("/agent/turn", body, timeout=180 if images else 90)
    if resp.get("error"):
        return {
            "error": resp["error"],
            "conversation_id": conversation_id,
            "tool_uses": [],
            "text": f"Could not reach agent: {resp['error']}",
            "done": True,
            "thinking": "",
            "stop_reason": "error",
            "usage": {},
        }
    return resp


def post_results(
    conversation_id: str,
    results: list[dict],
    context: dict,
    model: Optional[str] = None,
) -> dict:
    """Post tool_result blocks for the previous tool_uses.

    `results` is a list of {"tool_use_id": str, "content": str, "is_error": bool}.
    Returns the next turn response (same shape as start_turn).
    """
    body = {
        "conversation_id": conversation_id,
        "results": results,
        "context": context,
    }
    if model:
        body["model"] = model
    resp = _post_json("/agent/result", body, timeout=120)
    if resp.get("error"):
        return {
            "error": resp["error"],
            "conversation_id": conversation_id,
            "tool_uses": [],
            "text": f"Could not reach agent: {resp['error']}",
            "done": True,
            "thinking": "",
            "stop_reason": "error",
            "usage": {},
        }
    return resp


# Map agent_v2 tool names → executor action types. Most are 1:1; a few
# rename for clarity in the v2 schema while keeping the executor stable.
_TOOL_TO_ACTION = {
    "web_open": "web_search",  # v2 renamed (executor's web_search opens browser)
}


def tool_use_to_action(tu: dict):
    """Convert a tool_use dict {id, name, input} into an executor.Action."""
    from executor import Action, Risk
    name = tu.get("name", "")
    inp = tu.get("input", {}) or {}
    action_type = _TOOL_TO_ACTION.get(name, name)
    risk = Risk(_TOOL_RISK.get(name, "yellow"))
    # The executor parses command as either raw text or JSON. We always send
    # JSON so payload fields stay structured.
    cmd_str = json.dumps(inp) if inp else action_type
    return Action(
        type=action_type,
        description=_describe_tool_use(name, inp),
        command=cmd_str,
        risk=risk,
    )


def _describe_tool_use(name: str, inp: dict) -> str:
    """Human-readable description for a tool_use (shown in panel + logs)."""
    if name == "search_files":
        q = inp.get("query", "?")
        ft = inp.get("file_type", "")
        return f"Search {ft or 'files'} for '{q}'"
    if name == "open_file":
        return f"Open {Path(inp.get('path', '?')).name}"
    if name == "open_app":
        return f"Open {inp.get('name', '?')}"
    if name == "open_url":
        return f"Open {inp.get('url', '?')[:60]}"
    if name == "read_file":
        return f"Read {Path(inp.get('path', '?')).name}"
    if name == "write_file":
        return f"Write {Path(inp.get('path', '?')).name}"
    if name == "applescript":
        script = inp.get("script", "")
        # Pull the first meaningful line
        for line in script.split("\n"):
            line = line.strip()
            if line and not line.startswith("--"):
                return f"AppleScript: {line[:60]}"
        return "AppleScript"
    if name == "shell":
        return f"Shell: {inp.get('command', '?')[:60]}"
    if name == "clipboard_copy":
        return "Copy to clipboard"
    if name == "clipboard_read":
        return "Read clipboard"
    if name == "notify":
        return f"Notify: {inp.get('message', '?')[:60]}"
    if name == "data_export":
        return f"Export {inp.get('source_app', '?')} → {inp.get('destination', '?')}"
    if name == "play_media":
        return f"Play '{inp.get('query', '?')}' on {inp.get('platform', '?')}"
    if name == "web_open":
        return f"Open {inp.get('engine', 'google')} search: '{inp.get('query', '?')}'"
    if name == "fetch_url":
        return f"Fetch {inp.get('url', '?')[:60]}"
    if name == "check_inbox":
        return f"Check inbox ({inp.get('max_results', 10)} emails)"
    if name == "search_email":
        return f"Search email: '{inp.get('query', '?')}'"
    if name == "read_email":
        return f"Read email"
    if name == "send_email":
        return f"Send email to {inp.get('to', '?')}"
    if name == "draft_email":
        return f"Draft email to {inp.get('to', '?')}"
    if name == "screenshot":
        return "Capture screen"
    if name == "click_at":
        return f"Click at ({inp.get('x', '?')}, {inp.get('y', '?')})"
    if name == "type_text":
        t = inp.get("text", "")
        return f"Type: {t[:40]}" + ("…" if len(t) > 40 else "")
    if name == "key_combo":
        return f"Press {inp.get('keys', '?')}"
    if name == "scroll":
        return f"Scroll {inp.get('direction', 'down')}"
    if name == "wait":
        return f"Wait {inp.get('seconds', 1)}s"
    if name == "save_memory":
        return f"Remember: {inp.get('fact', '')[:50]}"
    if name == "set_shortcut":
        return f"Create shortcut /{inp.get('trigger', '?')}"
    return name


def execute_tool_use(tu: dict) -> dict:
    """Run a single tool_use through the executor. Returns a tool_result dict
    suitable for posting back to /agent/result.

    Output content is capped to keep the conversation small.
    """
    from executor import execute_action

    action = tool_use_to_action(tu)
    sys.stderr.write(f"[agent_v2] executing tool_use: {action.type}({tu.get('input', {})})\n")
    executed = execute_action(action)

    # Build the content string for the model
    if executed.success:
        body = (executed.result or "(no output)")
    else:
        body = f"ERROR: {executed.error or executed.result or 'unknown'}"

    # Cap content to ~8K chars per tool result to keep convo small
    if len(body) > 8000:
        body = body[:8000] + f"\n... [truncated; full output was {len(body):,} chars]"

    return {
        "tool_use_id": tu["id"],
        "content": body,
        "is_error": not executed.success,
        # Local-only metadata for the UI:
        "_action": executed,
    }


def run_agent_loop(
    user_message: str,
    context: dict,
    images: Optional[list[dict]] = None,
    max_steps: int = 5,
    on_step=None,
) -> dict:
    """Drive the full agent loop until done or max_steps reached.

    `on_step(step_info)` is called after each round with:
      {
        "round": int,
        "tool_uses": [...],
        "executed": [Action, ...],
        "text": str,
        "thinking": str,
      }

    Returns a summary dict:
      {
        "conversation_id": str,
        "final_text": str,
        "rounds": int,
        "all_executed": [Action, ...],
        "error": str | None,
      }
    """
    all_executed: list = []
    conversation_id: Optional[str] = None
    final_text = ""
    last_thinking = ""
    total_cost = 0.0
    last_warning = ""
    last_today_total = 0.0
    last_model = ""

    # Round 1: start the conversation
    resp = start_turn(user_message, context, conversation_id, images)
    conversation_id = resp.get("conversation_id")

    for step in range(1, max_steps + 1):
        if resp.get("error"):
            return {
                "conversation_id": conversation_id,
                "final_text": resp.get("text", ""),
                "rounds": step - 1,
                "all_executed": all_executed,
                "error": resp["error"],
                "total_cost_usd": total_cost,
                "today_total_usd": last_today_total,
                "warning": last_warning,
                "model": last_model,
            }

        tool_uses = resp.get("tool_uses", [])
        text = resp.get("text", "")
        thinking = resp.get("thinking", "")
        last_thinking = thinking or last_thinking
        # Accumulate cost across rounds
        total_cost += float(resp.get("cost_usd", 0.0) or 0.0)
        last_today_total = float(resp.get("cost_today_usd", last_today_total) or last_today_total)
        last_warning = resp.get("cost_warning", "") or last_warning
        last_model = resp.get("model", "") or last_model

        # Budget-exceeded: server already returned an error message in `text`
        if resp.get("stop_reason") == "budget_exceeded":
            return {
                "conversation_id": conversation_id,
                "final_text": text,
                "rounds": step,
                "all_executed": all_executed,
                "error": "budget_exceeded",
                "total_cost_usd": total_cost,
                "today_total_usd": last_today_total,
                "warning": last_warning,
                "model": last_model,
            }

        # No tool calls → the agent is asking a question or finishing
        if not tool_uses:
            final_text = text
            if on_step:
                try:
                    on_step({
                        "round": step,
                        "tool_uses": [],
                        "executed": [],
                        "text": text,
                        "thinking": thinking,
                    })
                except Exception:
                    pass
            break

        # Execute each tool_use
        results = []
        executed_actions = []
        for tu in tool_uses:
            r = execute_tool_use(tu)
            executed_actions.append(r.pop("_action", None))
            results.append(r)
        all_executed.extend(a for a in executed_actions if a is not None)

        if on_step:
            try:
                on_step({
                    "round": step,
                    "tool_uses": tool_uses,
                    "executed": executed_actions,
                    "text": text,
                    "thinking": thinking,
                })
            except Exception:
                pass

        # Done if response said so
        if resp.get("done") and not tool_uses:
            final_text = text
            break

        # Post results, loop
        resp = post_results(conversation_id, results, context)

    return {
        "conversation_id": conversation_id,
        "final_text": final_text or resp.get("text", ""),
        "rounds": step,
        "all_executed": all_executed,
        "thinking": last_thinking,
        "error": resp.get("error"),
        "total_cost_usd": total_cost,
        "today_total_usd": last_today_total,
        "warning": last_warning,
        "model": last_model,
    }
