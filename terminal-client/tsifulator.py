#!/usr/bin/env python3
"""
tsifl — Terminal Client
Run: python3 terminal-client/tsifulator.py
"""

import os
import sys
import json
import subprocess
import readline
import urllib.request
import urllib.error
import ssl
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────

BACKEND_URL = os.getenv("TSIFULATOR_BACKEND_URL", "https://focused-solace-production-6839.up.railway.app")
CONFIG_PATH  = Path.home() / ".tsifulator_user"
HISTORY_FILE = Path.home() / ".tsifulator_history"

# SSL context — bypasses cert verification for Railway's endpoint
SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode    = ssl.CERT_NONE

# ── ANSI Colors — bright, clear Greek blue palette ───────────────────────────

RESET   = "\033[0m"
BOLD    = "\033[1m"
DIM     = "\033[2m"

# Greek flag blue — vivid
BLUE    = "\033[38;2;13;94;175m"
LBLUE   = "\033[38;2;66;153;225m"    # Lighter blue for secondary text
WHITE   = "\033[38;2;255;255;255m"   # Pure white for replies
MUTED   = "\033[38;2;140;160;185m"   # Soft blue-grey for chrome/borders
GREEN   = "\033[38;2;52;211;153m"    # Bright mint green for actions
YELLOW  = "\033[38;2;251;191;36m"    # Amber for warnings
RED     = "\033[38;2;248;113;113m"   # Soft red for errors

# Background highlights
BG_BLUE = "\033[48;2;13;94;175m"     # Solid blue background (for prompt badge)
BG_DARK = "\033[48;2;15;25;40m"      # Very dark blue-black for reply blocks

def b(s):   return f"{BLUE}{BOLD}{s}{RESET}"
def lb(s):  return f"{LBLUE}{s}{RESET}"
def w(s):   return f"{WHITE}{s}{RESET}"
def m(s):   return f"{MUTED}{s}{RESET}"
def g(s):   return f"{GREEN}{s}{RESET}"
def r(s):   return f"{RED}{s}{RESET}"
def y(s):   return f"{YELLOW}{s}{RESET}"
def dim(s): return f"{DIM}{MUTED}{s}{RESET}"

# ── User Identity ─────────────────────────────────────────────────────────────

def get_user_id():
    if CONFIG_PATH.exists():
        uid = CONFIG_PATH.read_text().strip()
        if uid:
            return uid
    env_id = os.environ.get("TSIFULATOR_USER_ID", "")
    if env_id:
        return env_id
    return "terminal-user-001"

# ── Context ───────────────────────────────────────────────────────────────────

def get_terminal_context():
    ctx = {
        "app":         "terminal",
        "shell":       os.environ.get("SHELL", "zsh"),
        "working_dir": os.getcwd(),
        "user":        os.environ.get("USER", ""),
        "os":          sys.platform,
    }
    try:
        hist_file = Path.home() / (
            ".zsh_history" if "zsh" in ctx["shell"] else ".bash_history"
        )
        if hist_file.exists():
            lines = hist_file.read_text(errors="replace").strip().splitlines()
            ctx["recent_commands"] = [l.lstrip(": 0123456789;") for l in lines[-5:]]
    except Exception:
        pass
    try:
        ctx["ls"] = os.listdir(".")[:20]
    except Exception:
        pass
    return ctx

# ── Backend ───────────────────────────────────────────────────────────────────

def send_message(user_id, message, context):
    payload = json.dumps({
        "user_id": user_id,
        "message": message,
        "context": context,
    }).encode("utf-8")

    req = urllib.request.Request(
        f"{BACKEND_URL}/chat/",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as resp:
        return json.loads(resp.read().decode("utf-8"))

# ── Actions ───────────────────────────────────────────────────────────────────

def execute_action(action):
    atype   = action.get("type", "")
    payload = action.get("payload", {})

    if atype == "run_shell_command":
        cmd = payload.get("command", "")
        if not cmd:
            return
        print(f"\n  {m('┌─')} {g('Running')}  {lb(cmd)}")
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=30)
            lines  = (result.stdout + result.stderr).strip().splitlines()
            for line in lines:
                print(f"  {m('│')}  {w(line)}")
            status = g("✓ Done") if result.returncode == 0 else r(f"✗ Exit {result.returncode}")
            print(f"  {m('└─')} {status}\n")
        except subprocess.TimeoutExpired:
            print(f"  {m('└─')} {y('⚠ Timed out (30s)')}\n")
        except Exception as e:
            print(f"  {m('└─')} {r(f'Error: {e}')}\n")

    elif atype == "write_file":
        path    = payload.get("path", "")
        content = payload.get("content", "")
        if path and content:
            try:
                Path(path).write_text(content)
                print(f"\n  {g('✓')} {m('Wrote')} {lb(path)}\n")
            except Exception as e:
                print(f"\n  {r(f'✗ {e}')}\n")

    elif atype == "open_url":
        url = payload.get("url", "")
        if url:
            subprocess.run(["open", url], check=False)
            print(f"\n  {g('✓')} {m('Opened')} {lb(url)}\n")

# ── Display ───────────────────────────────────────────────────────────────────

DIVIDER = m("  " + "─" * 48)

def print_header(user_id):
    uid_short = user_id[:20] + "..." if len(user_id) > 20 else user_id
    print()
    print(f"  {BLUE}{BOLD}⚡ tsifl{RESET}  {m('— Terminal')}")
    print(f"  {m('User')}  {lb(uid_short)}")
    print(f"  {m('Memory  Excel · RStudio · Terminal · Gmail')}")
    print(DIVIDER)
    print(f"  {dim('Type /exit to quit  ·  /clear to reset  ·  /user for ID')}")
    print()

def print_reply(reply, tasks_remaining):
    print()
    # Word-wrap at 68 chars
    words = reply.split()
    line  = ""
    lines = []
    for word in words:
        if len(line) + len(word) + 1 > 68:
            lines.append(line)
            line = word
        else:
            line = (line + " " + word).strip()
    if line:
        lines.append(line)

    print(f"  {BLUE}{BOLD}tsifl{RESET}")
    for l in lines:
        print(f"  {WHITE}{l}{RESET}")

    print()
    if tasks_remaining >= 0:
        print(f"  {dim(f'{tasks_remaining} tasks remaining')}")
    print()

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if HISTORY_FILE.exists():
        try:
            readline.read_history_file(str(HISTORY_FILE))
        except Exception:
            pass
    readline.set_history_length(500)

    user_id = get_user_id()
    print_header(user_id)

    try:
        while True:
            try:
                sys.stdout.write(f"  {BLUE}{BOLD}You{RESET}  {MUTED}›{RESET} ")
                sys.stdout.flush()
                user_input = input("").strip()
            except (EOFError, KeyboardInterrupt):
                print(f"\n  {m('Goodbye.')}\n")
                break

            if not user_input:
                continue

            if user_input.lower() in ("/exit", "/quit", "exit", "quit"):
                print(f"\n  {m('Goodbye.')}\n")
                break

            if user_input.lower() == "/clear":
                os.system("clear")
                print_header(user_id)
                continue

            if user_input.lower() == "/user":
                print(f"\n  {m('User ID')}  {lb(user_id)}\n")
                continue

            # Thinking indicator
            sys.stdout.write(f"  {MUTED}Thinking...{RESET}\n")
            sys.stdout.flush()

            try:
                context = get_terminal_context()
                data    = send_message(user_id, user_input, context)

                # Clear thinking line
                sys.stdout.write("\033[F\033[K")

                print_reply(data.get("reply", ""), data.get("tasks_remaining", -1))

                actions = data.get("actions", [])
                action  = data.get("action", {})
                if actions:
                    for a in actions:
                        execute_action(a)
                elif action and action.get("type", "none") != "none":
                    execute_action(action)

            except urllib.error.HTTPError as e:
                sys.stdout.write("\033[F\033[K")
                try:
                    err = json.loads(e.read())
                    print(f"\n  {r('✗')} {w(err.get('detail', str(e)))}\n")
                except Exception:
                    print(f"\n  {r(f'✗ HTTP {e.code}')}\n")

            except Exception as e:
                sys.stdout.write("\033[F\033[K")
                print(f"\n  {r('✗ Could not reach backend')}")
                print(f"  {dim(str(e))}\n")

    finally:
        try:
            readline.write_history_file(str(HISTORY_FILE))
        except Exception:
            pass

if __name__ == "__main__":
    main()
