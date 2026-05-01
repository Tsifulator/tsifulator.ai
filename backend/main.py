"""
Tsifulator.ai — Backend Brain
The central server that powers all tsifl integrations.
All chat, memory, auth, and actions route through here.
"""

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from dotenv import load_dotenv
import os
import time
import json
import logging
try:
    import psutil
except ImportError:
    psutil = None
from datetime import datetime
from pathlib import Path

# Load .env locally — in Railway, env vars come from the dashboard instead
env_path = Path(__file__).parent.parent / ".env"
if env_path.exists():
    load_dotenv(env_path, override=True)

# Structured JSON logging (Improvement 89)
logger = logging.getLogger("tsifl")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler()
handler.setFormatter(logging.Formatter('%(message)s'))
logger.addHandler(handler)

# Server stats (Improvement 90)
_server_start_time = time.time()
_total_requests = 0
_last_error_time = None

app = FastAPI(
    title="Tsifulator.ai API",
    description="Agentic Sandbox for Financial Analysts",
    version="0.3.0"
)

# CORS — allow Office add-ins, Chrome extensions, and production frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://localhost:3000",
        "https://localhost:3001",
        "https://localhost:3002",
        "http://localhost:3000",
        "http://localhost:3001",
        "http://localhost:3002",
        "http://localhost:8000",
        "https://focused-solace-production-6839.up.railway.app",
        "null",  # Office add-ins send Origin: null
    ],
    allow_origin_regex=r"^((chrome-extension|moz-extension|vscode-webview)://.*|https://([a-z0-9-]+\.)*(officeapps\.live\.com|office\.com|office365\.com|microsoft\.com|sharepoint\.com))$",
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Request logging middleware (Improvement 89)
@app.middleware("http")
async def log_requests(request: Request, call_next):
    global _total_requests, _last_error_time
    _total_requests += 1
    start = time.time()
    # NOTE: do NOT read request.body() after call_next — the ASGI stream is
    # consumed by the handler and a second read blocks the connection, causing
    # all POST requests to hang. Read user_id from query params only (safe).
    user_id = request.query_params.get("user_id", "")
    response = await call_next(request)
    latency_ms = round((time.time() - start) * 1000, 1)
    if response.status_code >= 500:
        _last_error_time = datetime.utcnow().isoformat()
    log_entry = {
        "ts": datetime.utcnow().isoformat(),
        "method": request.method,
        "path": request.url.path,
        "status": response.status_code,
        "latency_ms": latency_ms,
    }
    if user_id:
        log_entry["user_id"] = user_id
    logger.info(json.dumps(log_entry))
    return response

# Global error handler — never expose stack traces
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    global _last_error_time
    _last_error_time = datetime.utcnow().isoformat()
    # Always include the exception class + message in the response so we can
    # diagnose 500s in production. The stack trace is still kept out.
    return JSONResponse(
        status_code=500,
        content={
            "error":  "Internal server error",
            "detail": f"{type(exc).__name__}: {exc}"[:500],
        },
    )

# Startup: auto-create Supabase tables if missing (Improvement 1)
@app.on_event("startup")
async def startup_create_tables():
    try:
        url = os.getenv("SUPABASE_URL", "")
        key = os.getenv("SUPABASE_KEY", "")
        if not url or not key or url == "your_supabase_project_url":
            logger.info(json.dumps({"event": "startup", "supabase": "not_configured"}))
            return

        # Try to ensure the sessions table exists (critical for auth persistence)
        from routes.auth import ensure_sessions_table
        await ensure_sessions_table()

        from supabase import create_client
        sb = create_client(url, key)
        # Check other tables
        for table in ["notes", "usage"]:
            try:
                sb.table(table).select("*").limit(1).execute()
                logger.info(json.dumps({"event": "table_check_ok", "table": table}))
            except Exception:
                logger.info(json.dumps({"event": "table_check_failed", "table": table}))
    except Exception as e:
        logger.info(json.dumps({"event": "startup_table_check_error", "error": str(e)}))

@app.get("/")
def health_check():
    return {
        "status": "Tsifulator.ai is running",
        "version": "0.3.0",
        "env": os.getenv("ENV", "production")
    }

# Health dashboard (Improvement 90)
@app.get("/health")
def health_dashboard():
    from routes.auth import _session_store, _supabase_table_ok
    uptime_seconds = round(time.time() - _server_start_time)
    hours, remainder = divmod(uptime_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    # Check Supabase connection
    sb_status = "unknown"
    try:
        url = os.getenv("SUPABASE_URL", "")
        key = os.getenv("SUPABASE_KEY", "")
        if url and key and url != "your_supabase_project_url":
            from supabase import create_client
            sb = create_client(url, key)
            # Try messages table (always exists from Supabase auth)
            try:
                sb.table("sessions").select("id").limit(1).execute()
                sb_status = "connected"
            except Exception:
                # Table might not exist but Supabase itself is reachable
                sb_status = "connected_no_tables"
        else:
            sb_status = "not_configured"
    except Exception:
        sb_status = "disconnected"
    return {
        "status": "healthy",
        "uptime": f"{hours}h {minutes}m {seconds}s",
        "uptime_seconds": uptime_seconds,
        "total_requests": _total_requests,
        "active_sessions": len(_session_store),
        "memory_mb": round(psutil.Process().memory_info().rss / 1024 / 1024, 1) if psutil else None,
        "last_error": _last_error_time,
        "supabase": sb_status,
    }

# --- Routes ---
from routes import chat, auth, gmail, files, notes
from routes import transfer, calendar, computer_use

app.include_router(chat.router, prefix="/chat")
app.include_router(auth.router, prefix="/auth")
app.include_router(gmail.router, prefix="/gmail")
app.include_router(files.router, prefix="/files")
app.include_router(notes.router, prefix="/notes")
app.include_router(transfer.router, prefix="/transfer")
app.include_router(calendar.router, prefix="/calendar")
app.include_router(computer_use.router)

# --- Notes App (served as static HTML) ---
NOTES_APP_PATH = Path(__file__).parent / "static" / "notes.html"

@app.get("/notes-app")
async def serve_notes_app():
    if NOTES_APP_PATH.exists():
        return FileResponse(NOTES_APP_PATH, media_type="text/html")
    from fastapi.responses import HTMLResponse
    return HTMLResponse(
        content='<html><body style="font-family:sans-serif;text-align:center;padding:60px;">'
        '<h1 style="color:#0D5EAF;">tsifl Notes</h1>'
        '<p>Notes app file not found.</p></body></html>',
        status_code=200
    )

# --- Launch App Endpoint (for cross-app capability) ---
import subprocess
from pydantic import BaseModel

class LaunchAppRequest(BaseModel):
    app_name: str

APP_NAMES = {
    "microsoft excel": "Microsoft Excel",
    "excel": "Microsoft Excel",
    "microsoft powerpoint": "Microsoft PowerPoint",
    "powerpoint": "Microsoft PowerPoint",
    "microsoft word": "Microsoft Word",
    "word": "Microsoft Word",
    "visual studio code": "Visual Studio Code",
    "vscode": "Visual Studio Code",
    "vs code": "Visual Studio Code",
    "rstudio": "RStudio",
    "r studio": "RStudio",
    "notes": "Notes",
    "terminal": "Terminal",
    "safari": "Safari",
    "chrome": "Google Chrome",
    "finder": "Finder",
    "calendar": "Calendar",
}

@app.post("/launch-app")
async def launch_app(request: LaunchAppRequest):
    import platform
    app_key = request.app_name.lower().strip()

    # Special case: notes-app opens in browser
    if app_key in ("notes", "tsifl notes"):
        return {"status": "ok", "url": "https://focused-solace-production-6839.up.railway.app/notes-app", "message": "Open Notes in browser"}

    app_name = APP_NAMES.get(app_key)
    if not app_name:
        return {"status": "error", "message": f"Unknown app: {request.app_name}. Available: {', '.join(set(APP_NAMES.values()))}"}

    system = platform.system()
    if system == "Darwin":
        try:
            subprocess.Popen(["open", "-a", app_name])
            return {"status": "ok", "message": f"Launching {app_name}"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
    elif system == "Windows":
        try:
            subprocess.Popen(["start", "", app_name], shell=True)
            return {"status": "ok", "message": f"Launching {app_name}"}
        except Exception as e:
            return {"status": "error", "message": str(e)}
    else:
        return {"status": "info", "message": f"Launch requested for {app_name} — run locally to use this feature"}

@app.get("/debug/versions")
def debug_versions():
    import anthropic
    return {"anthropic": anthropic.__version__}

