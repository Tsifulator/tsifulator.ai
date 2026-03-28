"""
Gmail Route — read, search, draft, and send emails via the Gmail API.
Users authenticate once with Google OAuth. Token is stored at ~/.tsifulator_gmail_token.json.
"""

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from pathlib import Path
import json
import base64
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

router = APIRouter()

TOKEN_PATH = Path.home() / ".tsifulator_gmail_token.json"
CREDS_PATH = Path(__file__).parent.parent.parent / "gmail-client" / "credentials.json"

# ── Helper: build Gmail service ───────────────────────────────────────────────

def get_gmail_service():
    """Build an authenticated Gmail API service object."""
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except ImportError:
        raise HTTPException(
            status_code=503,
            detail="Gmail integration not installed. Run: pip install google-api-python-client google-auth-oauthlib"
        )

    if not TOKEN_PATH.exists():
        raise HTTPException(
            status_code=401,
            detail="Gmail not connected. Run: python3 gmail-client/gmail_setup.py"
        )

    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH))

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN_PATH.write_text(creds.to_json())

    return build("gmail", "v1", credentials=creds)

# ── Models ────────────────────────────────────────────────────────────────────

class EmailRequest(BaseModel):
    to: str
    subject: str
    body: str
    reply_to_id: str = ""   # Gmail message ID to reply to

class SearchRequest(BaseModel):
    query: str
    max_results: int = 10

# ── Endpoints ─────────────────────────────────────────────────────────────────

@router.get("/status")
async def gmail_status():
    """Check whether Gmail is connected."""
    connected = TOKEN_PATH.exists()
    return {
        "connected": connected,
        "setup_command": "python3 gmail-client/gmail_setup.py" if not connected else None
    }

@router.get("/inbox")
async def get_inbox(max_results: int = 20):
    """Fetch recent inbox messages."""
    service = get_gmail_service()
    try:
        results = service.users().messages().list(
            userId="me",
            labelIds=["INBOX"],
            maxResults=max_results
        ).execute()

        messages = results.get("messages", [])
        emails = []

        for msg in messages:
            detail = service.users().messages().get(
                userId="me",
                id=msg["id"],
                format="metadata",
                metadataHeaders=["From", "Subject", "Date"]
            ).execute()

            headers = {h["name"]: h["value"] for h in detail["payload"]["headers"]}
            emails.append({
                "id":      msg["id"],
                "from":    headers.get("From", ""),
                "subject": headers.get("Subject", "(no subject)"),
                "date":    headers.get("Date", ""),
                "snippet": detail.get("snippet", ""),
                "unread":  "UNREAD" in detail.get("labelIds", []),
            })

        return {"emails": emails, "count": len(emails)}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/message/{message_id}")
async def get_message(message_id: str):
    """Fetch the full body of a specific email."""
    service = get_gmail_service()
    try:
        msg = service.users().messages().get(
            userId="me", id=message_id, format="full"
        ).execute()

        headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}

        # Extract body text
        body = _extract_body(msg["payload"])

        return {
            "id":      message_id,
            "from":    headers.get("From", ""),
            "to":      headers.get("To", ""),
            "subject": headers.get("Subject", ""),
            "date":    headers.get("Date", ""),
            "body":    body,
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/search")
async def search_emails(req: SearchRequest):
    """Search emails using Gmail query syntax."""
    service = get_gmail_service()
    try:
        results = service.users().messages().list(
            userId="me", q=req.query, maxResults=req.max_results
        ).execute()

        messages = results.get("messages", [])
        emails = []

        for msg in messages:
            detail = service.users().messages().get(
                userId="me",
                id=msg["id"],
                format="metadata",
                metadataHeaders=["From", "Subject", "Date"]
            ).execute()
            headers = {h["name"]: h["value"] for h in detail["payload"]["headers"]}
            emails.append({
                "id":      msg["id"],
                "from":    headers.get("From", ""),
                "subject": headers.get("Subject", "(no subject)"),
                "date":    headers.get("Date", ""),
                "snippet": detail.get("snippet", ""),
            })

        return {"emails": emails, "count": len(emails), "query": req.query}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/send")
async def send_email(req: EmailRequest):
    """Send or reply to an email."""
    service = get_gmail_service()
    try:
        msg = MIMEMultipart("alternative")
        msg["To"]      = req.to
        msg["Subject"] = req.subject
        msg.attach(MIMEText(req.body, "plain"))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        body = {"raw": raw}

        if req.reply_to_id:
            body["threadId"] = req.reply_to_id

        sent = service.users().messages().send(userId="me", body=body).execute()
        return {"status": "sent", "id": sent["id"]}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/draft")
async def create_draft(req: EmailRequest):
    """Save a draft email (don't send yet)."""
    service = get_gmail_service()
    try:
        msg = MIMEMultipart("alternative")
        msg["To"]      = req.to
        msg["Subject"] = req.subject
        msg.attach(MIMEText(req.body, "plain"))

        raw   = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        draft = service.users().drafts().create(
            userId="me", body={"message": {"raw": raw}}
        ).execute()

        return {"status": "drafted", "id": draft["id"]}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ── Helper: extract email body ────────────────────────────────────────────────

def _extract_body(payload: dict) -> str:
    """Recursively extract plain text body from Gmail message payload."""
    if payload.get("mimeType") == "text/plain":
        data = payload.get("body", {}).get("data", "")
        if data:
            return base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="replace")

    if "parts" in payload:
        for part in payload["parts"]:
            result = _extract_body(part)
            if result:
                return result

    return ""
