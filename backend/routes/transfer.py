"""
Transfer Route — Cross-app data transfer (e.g. R plots to Excel).
Stores data with persistent JSON backing and configurable TTL.
"""

import json
import uuid
import time
import os
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional

router = APIRouter()

# Persistent store path
_STORE_FILE = "/tmp/.tsifl_transfer_store.json"

# In-memory transfer store (synced to disk)
_transfer_store: dict = {}

# TTL defaults
TTL_SECONDS = 300  # 5 minutes
LONG_TTL_SECONDS = 3600  # 1 hour for r_output and data_snapshot
_LONG_TTL_TYPES = {"r_output", "data_snapshot", "image", "ppt_actions"}


def _load_store():
    """Load store from disk on startup."""
    global _transfer_store
    if os.path.exists(_STORE_FILE):
        try:
            with open(_STORE_FILE, "r") as f:
                _transfer_store = json.load(f)
        except (json.JSONDecodeError, IOError):
            _transfer_store = {}


def _save_store():
    """Persist store to disk."""
    try:
        with open(_STORE_FILE, "w") as f:
            json.dump(_transfer_store, f)
    except IOError:
        pass


# Load on module import
_load_store()


def _get_ttl(data_type: str) -> int:
    """Return TTL based on data type."""
    return LONG_TTL_SECONDS if data_type in _LONG_TTL_TYPES else TTL_SECONDS


def _cleanup_expired():
    """Remove expired transfers."""
    now = time.time()
    expired = [
        k for k, v in _transfer_store.items()
        if now - v["created_at"] > _get_ttl(v.get("data_type", ""))
    ]
    if expired:
        for k in expired:
            del _transfer_store[k]
        _save_store()


class TransferStore(BaseModel):
    from_app: str
    to_app: str
    data_type: str  # "image", "table", "text", "r_output", "data_snapshot"
    data: str  # base64 for images, JSON string for tables, plain text
    metadata: Optional[dict] = {}


@router.post("/store")
async def store_transfer(request: TransferStore):
    """Store data for cross-app transfer. Returns a transfer_id."""
    _cleanup_expired()
    transfer_id = str(uuid.uuid4())[:8]
    _transfer_store[transfer_id] = {
        "from_app": request.from_app,
        "to_app": request.to_app,
        "data_type": request.data_type,
        "data": request.data,
        "metadata": request.metadata or {},
        "created_at": time.time(),
    }
    _save_store()
    ttl = _get_ttl(request.data_type)
    return {"transfer_id": transfer_id, "expires_in": ttl}


@router.get("/{transfer_id}")
async def get_transfer(transfer_id: str):
    """Retrieve and delete transfer data (one-time use)."""
    _cleanup_expired()
    item = _transfer_store.pop(transfer_id, None)
    if not item:
        raise HTTPException(status_code=404, detail="Transfer not found or expired")
    _save_store()
    return {
        "from_app": item["from_app"],
        "to_app": item["to_app"],
        "data_type": item["data_type"],
        "data": item["data"],
        "metadata": item["metadata"],
    }


@router.get("/pending/{to_app}")
async def check_pending(to_app: str):
    """Check if there are pending transfers for an app. Also includes broadcast items (to_app='any')."""
    _cleanup_expired()
    pending = [
        {"transfer_id": k, "from_app": v["from_app"], "data_type": v["data_type"], "metadata": v["metadata"]}
        for k, v in _transfer_store.items()
        if v["to_app"] == to_app or v["to_app"] == "any"
    ]
    return {"pending": pending}


@router.get("/context/{app}")
async def get_context_endpoint(app: str):
    """Return the latest 5 cross-app items from OTHER apps (read-only, non-consuming)."""
    _cleanup_expired()
    items = []
    for k, v in _transfer_store.items():
        if v["from_app"] != app:
            items.append({
                "transfer_id": k,
                "from_app": v["from_app"],
                "to_app": v["to_app"],
                "data_type": v["data_type"],
                "data": v["data"][:500] if len(v.get("data", "")) > 500 else v.get("data", ""),
                "metadata": v.get("metadata", {}),
                "created_at": v["created_at"],
            })
    # Sort by created_at descending, take last 5
    items.sort(key=lambda x: x["created_at"], reverse=True)
    return {"context": items[:5]}


def get_cross_app_context(app: str) -> str:
    """
    Helper function (not a route) that returns a formatted string of recent
    cross-app data from other apps. Used by chat.py to inject context.
    """
    _cleanup_expired()
    items = []
    now = time.time()
    for k, v in _transfer_store.items():
        if v["from_app"] != app:
            ttl = _get_ttl(v.get("data_type", ""))
            if now - v["created_at"] <= ttl:
                items.append(v)

    if not items:
        return ""

    # Sort by created_at descending, take last 5
    items.sort(key=lambda x: x["created_at"], reverse=True)
    items = items[:5]

    parts = []
    for item in items:
        data_preview = item.get("data", "")[:300]
        meta = item.get("metadata", {})
        meta_str = ", ".join(f"{k}={v}" for k, v in meta.items()) if meta else ""
        age = int(now - item["created_at"])
        parts.append(
            f"From {item['from_app']} ({age}s ago) type={item['data_type']}: "
            f"{data_preview}"
            + (f" | meta: {meta_str}" if meta_str else "")
        )

    return " | ".join(parts)
