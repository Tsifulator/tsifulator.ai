"""
Chat Route — the main entry point for all user messages.
Receives a message from any tsifl integration, pulls session-scoped history,
sends to Claude, saves response, returns action(s).
"""

import hashlib
import time
import base64
import os
from collections import OrderedDict
from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from services.claude import get_claude_response, get_claude_stream
from services.usage import check_and_increment_usage
from services.memory import save_message, get_recent_history, is_connected

# File extensions that should be saved to /tmp/ for import_csv
_SAVEABLE_EXTENSIONS = {".csv", ".tsv", ".txt", ".json", ".xml"}

router = APIRouter()

# Session-scoped conversation history (in-memory, LRU eviction)
MAX_SESSIONS = 50
MAX_MESSAGES_PER_SESSION = 10
_history_store: OrderedDict = OrderedDict()

# Response cache with TTL (Improvement 91)
_response_cache: OrderedDict = OrderedDict()
MAX_CACHE_ENTRIES = 100
CACHE_TTL_SECONDS = 300  # 5 minutes


def _cache_key(user_id: str, message: str, app: str) -> str:
    raw = f"{user_id}:{message}:{app}"
    return hashlib.md5(raw.encode()).hexdigest()


def _get_cached_response(key: str) -> dict | None:
    if key in _response_cache:
        entry = _response_cache[key]
        if time.time() - entry["ts"] < CACHE_TTL_SECONDS:
            _response_cache.move_to_end(key)
            return entry["data"]
        else:
            del _response_cache[key]
    return None


def _set_cached_response(key: str, data: dict):
    if len(_response_cache) >= MAX_CACHE_ENTRIES:
        _response_cache.popitem(last=False)
    _response_cache[key] = {"data": data, "ts": time.time()}


def _get_session_history(session_id: str) -> list:
    """Get conversation history for a session."""
    if not session_id:
        return []
    if session_id in _history_store:
        _history_store.move_to_end(session_id)
        return _history_store[session_id]
    return []


def _add_to_history(session_id: str, role: str, content: str):
    """Add a message to session history with LRU eviction."""
    if not session_id:
        return
    if session_id not in _history_store:
        # Evict oldest session if at capacity
        if len(_history_store) >= MAX_SESSIONS:
            _history_store.popitem(last=False)
        _history_store[session_id] = []
    _history_store.move_to_end(session_id)
    _history_store[session_id].append({"role": role, "content": content})
    # Keep only last N messages per session
    if len(_history_store[session_id]) > MAX_MESSAGES_PER_SESSION:
        _history_store[session_id] = _history_store[session_id][-MAX_MESSAGES_PER_SESSION:]


class ImageData(BaseModel):
    media_type: str = "image/png"
    data: str  # base64-encoded image/document data
    file_name: str = ""  # original filename (for documents)

class ChatRequest(BaseModel):
    user_id: str
    message: str
    context: dict = {}
    session_id: str = ""
    images: list[ImageData] = []

class ChatResponse(BaseModel):
    reply: str
    action: dict = {}
    actions: list = []
    tasks_remaining: int = -1
    memory_active: bool = False
    model_used: str = ""

@router.get("/debug")
async def debug():
    """Quick test: does the tool call work and return actions?"""
    from services.claude import get_claude_response
    result = await get_claude_response(
        message    = "Write the word Test in cell A1",
        context    = {"app": "excel", "sheet": "Sheet1", "sheet_data": []},
        session_id = "debug",
        history    = []
    )
    return {
        "reply":        result.get("reply"),
        "action":       result.get("action"),
        "actions":      result.get("actions"),
        "action_count": len(result.get("actions", [])) + (1 if result.get("action") else 0)
    }

@router.post("/", response_model=ChatResponse)
async def chat(request: ChatRequest):
    # 1. Check usage limit
    usage = await check_and_increment_usage(request.user_id)
    if not usage["allowed"]:
        raise HTTPException(
            status_code=429,
            detail="Monthly task limit reached. Upgrade to Pro for unlimited tasks."
        )

    # 2. Check cache for identical recent query (Improvement 91)
    app = request.context.get("app", "excel")
    cache_k = _cache_key(request.user_id, request.message, app)
    if not request.images:
        cached = _get_cached_response(cache_k)
        if cached:
            return ChatResponse(
                reply=cached["reply"],
                action=cached.get("action", {}),
                actions=cached.get("actions", []),
                tasks_remaining=usage["remaining"],
                memory_active=is_connected()
            )

    # 3. Get session-scoped history (last 10 messages for this session)
    # For heavy action apps (excel, rstudio, powerpoint), skip history
    # to avoid teaching Claude to return abbreviated actions.
    heavy_action_apps = {"excel", "rstudio", "powerpoint", "google_sheets"}
    if app in heavy_action_apps:
        history = []
    else:
        history = _get_session_history(request.session_id)

    # 4. Save the user's message
    _add_to_history(request.session_id, "user", request.message)
    await save_message(
        user_id=request.user_id,
        role="user",
        content=request.message,
        app=app,
        session_id=request.session_id
    )

    # 5. Save uploaded data files (CSV, TSV, etc.) to /tmp/ so import_csv can use them
    images = [{"media_type": img.media_type, "data": img.data, "file_name": img.file_name} for img in request.images] if request.images else []
    message = request.message
    saved_file_paths = []
    remaining_images = []
    for img in images:
        file_name = img.get("file_name", "")
        ext = ("." + file_name.rsplit(".", 1)[-1].lower()) if "." in file_name else ""
        if ext in _SAVEABLE_EXTENSIONS and img.get("data"):
            # Save to /tmp/ for import_csv
            safe_name = file_name.replace("/", "_").replace(" ", "_")
            save_path = f"/tmp/{safe_name}"
            try:
                raw = base64.b64decode(img["data"])
                with open(save_path, "wb") as f:
                    f.write(raw)
                saved_file_paths.append(save_path)
            except Exception:
                remaining_images.append(img)
        else:
            remaining_images.append(img)

    # Inject saved file paths into the message so Claude uses import_csv
    if saved_file_paths and app in {"excel", "google_sheets"}:
        paths_str = ", ".join(saved_file_paths)
        message = f"{message}\n\n[SYSTEM: The user uploaded data files that have been saved to the server. Use import_csv to import them into the spreadsheet. File paths: {paths_str}]"

    result = await get_claude_response(
        message=message,
        context=request.context,
        session_id=request.session_id,
        history=history,
        images=remaining_images
    )

    # 6. Save Claude's reply to history and persistent memory
    _add_to_history(request.session_id, "assistant", result["reply"])
    await save_message(
        user_id=request.user_id,
        role="assistant",
        content=result["reply"],
        app=app,
        session_id=request.session_id
    )

    # 7. Cache the response (Improvement 91)
    if not request.images:
        _set_cached_response(cache_k, result)

    return ChatResponse(
        reply=result["reply"],
        action=result.get("action", {}),
        actions=result.get("actions", []),
        tasks_remaining=usage["remaining"],
        memory_active=is_connected(),
        model_used=result.get("model_used", "")
    )


# Streaming endpoint (Improvement 92)
@router.post("/stream")
async def chat_stream(request: ChatRequest):
    """Stream Claude's text response as Server-Sent Events."""
    usage = await check_and_increment_usage(request.user_id)
    if not usage["allowed"]:
        raise HTTPException(
            status_code=429,
            detail="Monthly task limit reached. Upgrade to Pro for unlimited tasks."
        )

    app = request.context.get("app", "excel")
    heavy_action_apps = {"excel", "rstudio", "powerpoint", "google_sheets"}
    history = [] if app in heavy_action_apps else _get_session_history(request.session_id)

    _add_to_history(request.session_id, "user", request.message)
    await save_message(user_id=request.user_id, role="user", content=request.message, app=app, session_id=request.session_id)

    images = [{"media_type": img.media_type, "data": img.data, "file_name": img.file_name} for img in request.images] if request.images else []

    async def event_generator():
        full_reply = ""
        async for chunk in get_claude_stream(
            message=request.message,
            context=request.context,
            session_id=request.session_id,
            history=history,
            images=images
        ):
            full_reply += chunk
            yield f"data: {chunk}\n\n"
        yield "data: [DONE]\n\n"
        # Save after streaming complete
        _add_to_history(request.session_id, "assistant", full_reply)
        await save_message(user_id=request.user_id, role="assistant", content=full_reply, app=app, session_id=request.session_id)

    return StreamingResponse(event_generator(), media_type="text/event-stream")
