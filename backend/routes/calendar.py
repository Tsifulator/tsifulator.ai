"""
Calendar Route — Calendar event management.
Proxies to Google Calendar API or manages events locally.
For Google Workspace add-on, events are managed directly via Apps Script CalendarApp.
This provides a REST API for other integrations.
"""

import uuid
from datetime import datetime, timedelta
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional, List

router = APIRouter()

# In-memory event store (replace with Google Calendar API integration later)
_events_store = {}


class EventCreate(BaseModel):
    user_id: str
    title: str
    start_time: str  # ISO format
    end_time: Optional[str] = ""
    description: Optional[str] = ""
    attendees: Optional[List[str]] = []


class EventUpdate(BaseModel):
    title: Optional[str] = None
    start_time: Optional[str] = None
    end_time: Optional[str] = None
    description: Optional[str] = None
    attendees: Optional[List[str]] = None


@router.post("/events")
async def create_event(event: EventCreate):
    """Create a calendar event."""
    event_id = str(uuid.uuid4())[:8]
    now = datetime.utcnow().isoformat()
    data = {
        "id": event_id,
        "user_id": event.user_id,
        "title": event.title,
        "start_time": event.start_time,
        "end_time": event.end_time or "",
        "description": event.description or "",
        "attendees": event.attendees or [],
        "created_at": now,
    }
    if event.user_id not in _events_store:
        _events_store[event.user_id] = []
    _events_store[event.user_id].append(data)
    return data


@router.get("/events")
async def list_events(user_id: str, date: Optional[str] = None, days_ahead: Optional[int] = 7):
    """List upcoming events for a user."""
    events = _events_store.get(user_id, [])
    if date:
        try:
            target = datetime.fromisoformat(date)
            end = target + timedelta(days=days_ahead or 7)
            events = [
                e for e in events
                if target.isoformat() <= e["start_time"] <= end.isoformat()
            ]
        except Exception:
            pass
    # Sort by start time
    events.sort(key=lambda e: e.get("start_time", ""))
    return {"events": events}


@router.put("/events/{event_id}")
async def update_event(event_id: str, user_id: str, update: EventUpdate):
    """Update a calendar event."""
    events = _events_store.get(user_id, [])
    for event in events:
        if event["id"] == event_id:
            changes = {k: v for k, v in update.dict().items() if v is not None}
            event.update(changes)
            return event
    raise HTTPException(status_code=404, detail="Event not found")


@router.delete("/events/{event_id}")
async def delete_event(event_id: str, user_id: str):
    """Delete a calendar event."""
    events = _events_store.get(user_id, [])
    for i, event in enumerate(events):
        if event["id"] == event_id:
            events.pop(i)
            return {"status": "deleted"}
    raise HTTPException(status_code=404, detail="Event not found")
