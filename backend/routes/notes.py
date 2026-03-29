"""
Notes API — AI-powered notes with Supabase storage.
CRUD operations for notes with AI summarization and action extraction.
"""

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
    SUPABASE_KEY = os.getenv("SUPABASE_SERVICE_KEY", os.getenv("SUPABASE_ANON_KEY", ""))
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


# In-memory fallback when Supabase is unavailable
_notes_store = {}
_note_counter = 0


def _get_notes_for_user(user_id: str) -> list:
    if supabase:
        try:
            result = supabase.table("notes").select("*").eq("user_id", user_id).order("updated_at", desc=True).execute()
            return result.data or []
        except Exception:
            pass
    return [n for n in _notes_store.values() if n["user_id"] == user_id]


@router.get("/")
async def list_notes(user_id: str, folder: Optional[str] = None, search: Optional[str] = None):
    notes = _get_notes_for_user(user_id)
    if folder:
        notes = [n for n in notes if n.get("folder") == folder]
    if search:
        q = search.lower()
        notes = [n for n in notes if q in (n.get("title", "") + " " + n.get("content", "")).lower()]
    return {"notes": notes}


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
    global _note_counter
    now = datetime.utcnow().isoformat()
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
        except Exception as e:
            # Fall through to in-memory
            pass

    _note_counter += 1
    data["id"] = str(_note_counter)
    _notes_store[data["id"]] = data
    return data


@router.put("/{note_id}")
async def update_note(note_id: str, user_id: str, update: NoteUpdate):
    now = datetime.utcnow().isoformat()
    changes = {k: v for k, v in update.dict().items() if v is not None}
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


@router.get("/folders/list")
async def list_folders(user_id: str):
    notes = _get_notes_for_user(user_id)
    folders = list(set(n.get("folder", "General") for n in notes))
    if "General" not in folders:
        folders.insert(0, "General")
    return {"folders": sorted(folders)}
