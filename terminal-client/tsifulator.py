#!/usr/bin/env python3
"""
Tsifulator.ai — Terminal Client
Run this in any terminal to get the same AI brain as Excel + RStudio.
Same shared memory. Same user ID. Same Claude.

Usage:
    python3 tsifulator.py
"""

import os
import sys
import json
import subprocess
import readline
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────

BACKEND_URL = "https://focused-solace-production-6839.up.railway.app"
CONFIG_PATH = Path.home() / ".tsifulator_user"
HISTORY_FILE = Path.home() / ".tsifulator_history"

# ── ANSI Colors (Greek flag blue palette) ─────────────────────────────────────

class C:
    RESET   = "\033[0m"
    BOLD    = "\033[1m"
    DIM     = "\033[2m"
    BLUE    = "\033[38;2;13;94;175m"      # #0D5EAF — Greek flag blue
    LBLUE   = "\033[38;2;56;139;218m"     # Lighter blue for accents
    GREEN   = "\033[38;2;34;197;94m"      # #22c55e
    RED     = "\033[38;2;239;68;68m"      # #ef4444
    MUTED   = "\033[38;2;74;96;128m"      # Muted text
    WHITE   = "\033[38;2;226;232;240m"    # #e2e8f0
    BG_BLUE = "\033[48;2;13;94;175;22m"   # Subtle blue background

def blue(s):   return f"{C.BLUE}{C.BOLD}{s}{C.RESET}"
def green(s):  return f"{C.GREEN}{s}{C.RESET}"
def red(s):    return f"{C.RED}{s}{C.RESET}"
def muted(s):  return f"{C.MUTED}{s}{C.RESET}"
def white(s):  return f"{C.WHITE}{s}{C.RESET}"
def lblue(s):  return f"{C.LBLUE}{s}{C.RESET}"

# ── User Identity ─────────────────────────────────────────────────────────────

def get_user_id():
    """Read user ID — same one as Excel and RStudio use."""
    if CONFIG_PATH.exists():
        uid = CONFIG_PATH.read_text().strip()
        if uid:
            return uid
    env_id = os.environ.get("TSIFULATOR_USER_ID", "")
    if env_id:
        return env_id
    return "terminal-user-001"

# ── Terminal Context ──────────────────────────────────────────────────────────

def get_terminal_context():
    """Gather shell environment context to send with each message."""
    context = {
        "app": "terminal",
        "shell": os.environ.get("SHELL", "bash"),
        "working_dir": os.getcwd(),
        "user": os.environ.get("USER", ""),
        "os": sys.platform,
    }

    # Last few commands from shell history (bash/zsh)
    try:
        hist_file = Path.home() / (
            ".zsh_history" if "zsh" in context["shell"] else ".bash_history"
        )
        if hist_file.exists():
            lines = hist_file.read_text(errors="replace").strip().splitlines()
            recent = [l.lstrip(": 0123456789;") for l in lines[-5:]]
            context["recent_commands"] = recent
    except Exception:
        pass

    # Current directory listing
    try:
        files = os.listdir(".")[:20]
        context["ls"] = files
    except Exception:
        pass

    return context

# ── Backend Call ──────────────────────────────────────────────────────────────

def send_message(user_id: str, message: str, context: dict) -> dict:
    """POST to Tsifulator backend."""
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

    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))

# ── Action Execution ──────────────────────────────────────────────────────────

def execute_action(action: dict):
    """Execute a shell action returned by Claude."""
    action_type = action.get("type", "")
    payload     = action.get("payload", {})

    if action_type == "run_shell_command":
        cmd = payload.get("command", "")
        if not cmd:
            return

        print(f"\n{muted('─' * 50)}")
        print(f"{green('▶')} {muted('Running:')} {lblue(cmd)}")
        print(f"{muted('─' * 50)}")

        try:
            result = subprocess.run(
                cmd, shell=True, capture_output=True, text=True, timeout=30
            )
            if result.stdout.strip():
                print(white(result.stdout.strip()))
            if result.stderr.strip():
                print(red(result.stderr.strip()))
            if result.returncode == 0:
                print(f"\n{green('✅')} {muted('Done')}")
            else:
                print(f"\n{red('⚠')} {muted(f'Exit code: {result.returncode}')}")
        except subprocess.TimeoutExpired:
            print(red("⚠ Command timed out (30s limit)"))
        except Exception as e:
            print(red(f"⚠ Error: {e}"))

        print()

    elif action_type == "write_file":
        path    = payload.get("path", "")
        content = payload.get("content", "")
        if path and content:
            try:
                Path(path).write_text(content)
                print(f"{green('✅')} {muted('Wrote:')} {lblue(path)}\n")
            except Exception as e:
                print(red(f"⚠ Could not write {path}: {e}\n"))

    elif action_type == "open_url":
        url = payload.get("url", "")
        if url:
            subprocess.run(["open", url], check=False)
            print(f"{green('✅')} {muted('Opened:')} {lblue(url)}\n")

# ── Display ───────────────────────────────────────────────────────────────────

def print_header(user_id: str):
    print()
    print(f"  {blue('⚡ Tsifulator.ai')}  {muted('— Terminal')}")
    print(f"  {muted('User:')} {lblue(user_id[:16] + '...' if len(user_id) > 16 else user_id)}")
    print(f"  {muted('Shared memory: Excel + RStudio + Terminal')}")
    print(f"  {muted('Type your message. Type')} {lblue('/exit')} {muted('to quit.')}")
    print(f"  {muted('─' * 44)}")
    print()

def print_reply(reply: str, tasks_remaining: int):
    print()
    # Word-wrap the reply at ~70 chars for readability
    words = reply.split()
    line  = ""
    lines = []
    for word in words:
        if len(line) + len(word) + 1 > 72:
            lines.append(line)
            line = word
        else:
            line = (line + " " + word).strip()
    if line:
        lines.append(line)

    for l in lines:
        print(f"  {white(l)}")

    print()
    if tasks_remaining >= 0:
        print(f"  {muted(f'{tasks_remaining} tasks remaining')}")
    print()

# ── Main Loop ─────────────────────────────────────────────────────────────────

def main():
    # Set up readline history
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
                prompt = f"{blue('You')} {muted('›')} "
                # Use input() with ANSI in prompt
                sys.stdout.write(prompt)
                sys.stdout.flush()
                user_input = input("").strip()
            except (EOFError, KeyboardInterrupt):
                print(f"\n\n  {muted('Goodbye.')}\n")
                break

            if not user_input:
                continue

            # Slash commands
            if user_input.lower() in ("/exit", "/quit", "exit", "quit"):
                print(f"\n  {muted('Goodbye.')}\n")
                break

            if user_input.lower() == "/clear":
                os.system("clear")
                print_header(user_id)
                continue

            if user_input.lower() == "/user":
                print(f"\n  {muted('User ID:')} {lblue(user_id)}\n")
                continue

            # Send to backend
            print(f"  {muted('Thinking...')}")

            try:
                context = get_terminal_context()
                data    = send_message(user_id, user_input, context)

                # Clear "Thinking..." line
                sys.stdout.write("\033[F\033[K")

                print_reply(data.get("reply", ""), data.get("tasks_remaining", -1))

                # Execute actions
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
                    print(f"\n  {red('⚠')} {white(err.get('detail', str(e)))}\n")
                except Exception:
                    print(f"\n  {red('⚠')} {white(f'HTTP {e.code}')}\n")

            except Exception as e:
                sys.stdout.write("\033[F\033[K")
                print(f"\n  {red('⚠')} {white('Could not reach Tsifulator backend.')}")
                print(f"  {muted(str(e))}\n")

    finally:
        try:
            readline.write_history_file(str(HISTORY_FILE))
        except Exception:
            pass

if __name__ == "__main__":
    main()
