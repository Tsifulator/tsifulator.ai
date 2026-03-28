#!/usr/bin/env python3
"""
Tsifulator.ai — Gmail Setup
Run this ONCE to connect your Gmail account.
After this, Tsifulator can read, draft, and send emails from any app.

Usage:
    python3 gmail_setup.py

Requirements:
    pip install google-api-python-client google-auth-oauthlib
"""

import os
import sys
import json
from pathlib import Path

# ── ANSI Colors ───────────────────────────────────────────────────────────────

BLUE  = "\033[38;2;13;94;175m\033[1m"
GREEN = "\033[38;2;34;197;94m"
RED   = "\033[38;2;239;68;68m"
MUTED = "\033[38;2;74;96;128m"
RESET = "\033[0m"

TOKEN_PATH = Path.home() / ".tsifulator_gmail_token.json"
CREDS_PATH = Path(__file__).parent / "credentials.json"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.modify",
]

def main():
    print()
    print(f"  {BLUE}⚡ Tsifulator.ai — Gmail Setup{RESET}")
    print(f"  {MUTED}{'─' * 40}{RESET}")
    print()

    # Check dependencies
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
    except ImportError:
        print(f"  {RED}⚠ Missing dependencies. Run:{RESET}")
        print(f"  pip install google-api-python-client google-auth-oauthlib\n")
        sys.exit(1)

    # Check credentials.json
    if not CREDS_PATH.exists():
        print(f"  {RED}⚠ credentials.json not found.{RESET}")
        print()
        print(f"  {MUTED}To connect Gmail:{RESET}")
        print(f"  {MUTED}1. Go to: https://console.cloud.google.com{RESET}")
        print(f"  {MUTED}2. Create a project → Enable Gmail API{RESET}")
        print(f"  {MUTED}3. Credentials → Create OAuth 2.0 Client ID{RESET}")
        print(f"  {MUTED}   Type: Desktop app{RESET}")
        print(f"  {MUTED}4. Download JSON → save as gmail-client/credentials.json{RESET}")
        print()
        sys.exit(1)

    # Already authenticated?
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        if creds and creds.valid:
            service = build("gmail", "v1", credentials=creds)
            profile = service.users().getProfile(userId="me").execute()
            print(f"  {GREEN}✅ Already connected: {profile['emailAddress']}{RESET}")
            print(f"  {MUTED}Gmail integration is active.{RESET}\n")
            return

    # Run OAuth flow
    print(f"  {MUTED}Opening browser for Google login...{RESET}\n")
    flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_PATH), SCOPES)
    creds = flow.run_local_server(port=0, open_browser=True)

    # Save token
    TOKEN_PATH.write_text(creds.to_json())

    # Verify
    from googleapiclient.discovery import build
    service = build("gmail", "v1", credentials=creds)
    profile = service.users().getProfile(userId="me").execute()

    print()
    print(f"  {GREEN}✅ Connected: {profile['emailAddress']}{RESET}")
    print(f"  {MUTED}Token saved to ~/.tsifulator_gmail_token.json{RESET}")
    print()
    print(f"  {MUTED}Tsifulator can now read and send emails.{RESET}")
    print(f"  {MUTED}The token refreshes automatically.{RESET}\n")

if __name__ == "__main__":
    main()
