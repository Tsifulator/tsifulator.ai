"""
Computer Use Routes — endpoints for the hybrid architecture.
- Add-in polls /status/{id} to check when GUI tasks complete
- Desktop agent polls /pending for new tasks
- Desktop agent claims tasks via /claim/{id}
- Desktop agent reports results via /complete/{id}
"""

import asyncio
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from services.computer_use import (
    get_session,
    create_session,
    execute_session,
    _sessions,
)

router = APIRouter(prefix="/computer-use", tags=["computer-use"])


class SessionRequest(BaseModel):
    actions: list
    context: dict = {}


class SessionResponse(BaseModel):
    session_id: str
    status: str


class SessionStatus(BaseModel):
    session_id: str
    status: str  # pending, running, completed, failed
    result: dict | None = None
    error: str | None = None
    steps_taken: int = 0


class CompletionReport(BaseModel):
    status: str
    message: str = ""
    error: str = ""
    steps_taken: int = 0


@router.post("/start", response_model=SessionResponse)
async def start_session(request: SessionRequest):
    """Start a new computer use session. Returns immediately with a session_id.
    The session stays 'pending' until claimed by the desktop agent."""
    session_id = create_session(request.actions, request.context)
    # Don't auto-execute on server — wait for desktop agent to claim it
    return SessionResponse(session_id=session_id, status="pending")


@router.get("/status/{session_id}", response_model=SessionStatus)
async def check_status(session_id: str):
    """Check the status of a computer use session (polled by the add-in)."""
    session = get_session(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    return SessionStatus(
        session_id=session["id"],
        status=session["status"],
        result=session.get("result"),
        error=session.get("error"),
        steps_taken=len(session.get("steps", [])),
    )


@router.get("/pending")
async def get_pending_sessions():
    """Get all pending sessions (polled by the desktop agent)."""
    pending = [
        {
            "id": s["id"],
            "actions": s["actions"],
            "context": s["context"],
            "created_at": s["created_at"],
        }
        for s in _sessions.values()
        if s["status"] == "pending"
    ]
    return {"sessions": pending}


@router.post("/claim/{session_id}")
async def claim_session(session_id: str):
    """Desktop agent claims a session (marks it as running)."""
    session = get_session(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    if session["status"] != "pending":
        raise HTTPException(status_code=409, detail=f"Session is {session['status']}, not pending")
    session["status"] = "running"
    print(f"[computer_use] Session {session_id} claimed by desktop agent")
    return {"status": "running"}


@router.post("/cancel/{session_id}")
async def cancel_session(session_id: str):
    """Cancel a running or pending session (called by the add-in Stop button)."""
    session = get_session(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    if session["status"] in ("completed", "failed", "cancelled"):
        return {"status": session["status"], "message": "Already finished"}
    session["status"] = "cancelled"
    session["error"] = "Cancelled by user"
    print(f"[computer_use] Session {session_id} CANCELLED by user")
    return {"status": "cancelled"}


@router.post("/complete/{session_id}")
async def complete_session(session_id: str, report: CompletionReport):
    """Desktop agent reports completion of a session."""
    session = get_session(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Session not found")
    # Don't overwrite "cancelled" — the user already stopped this session
    if session["status"] == "cancelled":
        print(f"[computer_use] Session {session_id} already cancelled, ignoring completion report")
        return {"status": "cancelled"}
    session["status"] = report.status
    session["result"] = {
        "status": report.status,
        "message": report.message,
        "steps_taken": report.steps_taken,
    }
    if report.error:
        session["error"] = report.error
    print(f"[computer_use] Session {session_id} completed: {report.status} ({report.steps_taken} steps)")
    return {"status": "ok"}
