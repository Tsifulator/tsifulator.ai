"""
Notes API — AI-powered notes with Supabase storage.
CRUD operations for notes with AI summarization and action extraction.

Required Supabase table:
  CREATE TABLE IF NOT EXISTS notes (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    user_id TEXT NOT NULL,
    title TEXT NOT NULL DEFAULT 'Untitled',
    content TEXT DEFAULT '',
    folder TEXT DEFAULT 'General',
    tags TEXT[] DEFAULT '{}',
    created_at TIMESTAMPTZ DEFAULT now(),
    updated_at TIMESTAMPTZ DEFAULT now()
  );
  CREATE INDEX IF NOT EXISTS idx_notes_user ON notes(user_id);
"""

import uuid
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional, List
import os
from datetime import datetime

router = APIRouter()

# Supabase client
try:
    from supabase import create_client
    SUPABASE_URL = os.getenv("SUPABASE_URL", "https://dvynmzeyttwlmvunicqz.supabase.co")
    SUPABASE_KEY = os.getenv("SUPABASE_SERVICE_KEY", os.getenv("SUPABASE_ANON_KEY", os.getenv("SUPABASE_KEY", "")))
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_KEY else None
except Exception:
    supabase = None


class NoteCreate(BaseModel):
    user_id: str
    title: str
    content: str
    tags: Optional[List[str]] = []
    folder: Optional[str] = "General"


class NoteUpdate(BaseModel):
    title: Optional[str] = None
    content: Optional[str] = None
    tags: Optional[List[str]] = None
    folder: Optional[str] = None
    pinned: Optional[bool] = None
    user_id: Optional[str] = None


# In-memory fallback when Supabase is unavailable
_notes_store = {}


def _get_notes_for_user(user_id: str) -> list:
    if supabase:
        try:
            result = supabase.table("notes").select("*").eq("user_id", user_id).order("updated_at", desc=True).execute()
            return result.data or []
        except Exception:
            pass
    return sorted(
        [n for n in _notes_store.values() if n["user_id"] == user_id],
        key=lambda n: n.get("updated_at", ""),
        reverse=True
    )


@router.get("/")
async def list_notes(user_id: str, folder: Optional[str] = None, search: Optional[str] = None):
    notes = _get_notes_for_user(user_id)
    if folder:
        notes = [n for n in notes if n.get("folder") == folder]
    if search:
        q = search.lower()
        notes = [n for n in notes if q in (n.get("title", "") + " " + n.get("content", "")).lower()]
    return {"notes": notes}


@router.get("/folders/list")
async def list_folders(user_id: str):
    notes = _get_notes_for_user(user_id)
    folders = list(set(n.get("folder", "General") for n in notes))
    if "General" not in folders:
        folders.insert(0, "General")
    return {"folders": sorted(folders)}


@router.get("/{note_id}")
async def get_note(note_id: str, user_id: str):
    if supabase:
        try:
            result = supabase.table("notes").select("*").eq("id", note_id).eq("user_id", user_id).single().execute()
            return result.data
        except Exception:
            raise HTTPException(status_code=404, detail="Note not found")
    note = _notes_store.get(note_id)
    if not note or note["user_id"] != user_id:
        raise HTTPException(status_code=404, detail="Note not found")
    return note


@router.post("/")
async def create_note(note: NoteCreate):
    now = datetime.utcnow().isoformat()
    note_id = str(uuid.uuid4())
    data = {
        "user_id": note.user_id,
        "title": note.title,
        "content": note.content,
        "tags": note.tags or [],
        "folder": note.folder or "General",
        "created_at": now,
        "updated_at": now,
    }

    if supabase:
        try:
            result = supabase.table("notes").insert(data).execute()
            return result.data[0] if result.data else data
        except Exception:
            pass

    # In-memory fallback with proper UUID
    data["id"] = note_id
    _notes_store[note_id] = data
    return data


@router.put("/{note_id}")
async def update_note(note_id: str, user_id: str, update: NoteUpdate):
    now = datetime.utcnow().isoformat()
    changes = {k: v for k, v in update.dict().items() if v is not None and k != "user_id"}
    changes["updated_at"] = now

    if supabase:
        try:
            result = supabase.table("notes").update(changes).eq("id", note_id).eq("user_id", user_id).execute()
            return result.data[0] if result.data else changes
        except Exception:
            pass

    note = _notes_store.get(note_id)
    if not note or note["user_id"] != user_id:
        raise HTTPException(status_code=404, detail="Note not found")
    note.update(changes)
    return note


@router.delete("/{note_id}")
async def delete_note(note_id: str, user_id: str):
    if supabase:
        try:
            supabase.table("notes").delete().eq("id", note_id).eq("user_id", user_id).execute()
            return {"status": "deleted"}
        except Exception:
            pass

    if note_id in _notes_store and _notes_store[note_id]["user_id"] == user_id:
        del _notes_store[note_id]
        return {"status": "deleted"}
    raise HTTPException(status_code=404, detail="Note not found")


class NoteAIRequest(BaseModel):
    user_id: str
    action: str  # "summarize", "expand", "rewrite", "action_items", "ask"
    question: Optional[str] = ""  # For "ask" action


@router.post("/{note_id}/ai")
async def ai_action_on_note(note_id: str, request: NoteAIRequest):
    """Run AI on a note: summarize, expand, rewrite, extract action items, or ask a question."""
    note = None
    if supabase:
        try:
            result = supabase.table("notes").select("*").eq("id", note_id).eq("user_id", request.user_id).single().execute()
            note = result.data
        except Exception:
            pass
    if not note:
        note = _notes_store.get(note_id)
    if not note:
        raise HTTPException(status_code=404, detail="Note not found")

    prompts = {
        "summarize": f"Summarize this note concisely. Provide a TL;DR and key points as bullets.\n\nTitle: {note.get('title', '')}\n\n{note.get('content', '')}",
        "expand": f"Expand on this note with more detail, examples, and analysis. Keep the same structure but add depth.\n\nTitle: {note.get('title', '')}\n\n{note.get('content', '')}",
        "rewrite": f"Rewrite this note to be clearer, more professional, and better organized.\n\nTitle: {note.get('title', '')}\n\n{note.get('content', '')}",
        "action_items": f"Extract ALL action items, tasks, and to-dos from this note. Return as a numbered list with owner and deadline if mentioned.\n\nTitle: {note.get('title', '')}\n\n{note.get('content', '')}",
        "ask": f"{request.question}\n\nContext — Note title: {note.get('title', '')}\nNote content: {note.get('content', '')}",
    }

    prompt = prompts.get(request.action, prompts.get("ask", request.question or "Summarize this note."))

    try:
        from services.claude import get_claude_response
        result = await get_claude_response(
            message=prompt,
            context={"app": "notes", "note_title": note.get("title", ""), "note_content": note.get("content", "")},
            session_id=f"notes-ai-{note_id}",
        )
        return {"result": result.get("reply", "No response"), "action": request.action}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"AI error: {str(e)}")
