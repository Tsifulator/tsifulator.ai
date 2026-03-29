"""
Tsifulator.ai — Backend Brain
The central server that powers the Excel add-in, RStudio panel,
and all future integrations. All chat, memory, and actions route through here.
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from dotenv import load_dotenv
import os
from pathlib import Path

# Load .env locally — in Railway, env vars come from the dashboard instead
env_path = Path(__file__).parent.parent / ".env"
if env_path.exists():
    load_dotenv(env_path, override=True)

app = FastAPI(
    title="Tsifulator.ai API",
    description="Agentic Sandbox for Financial Analysts",
    version="0.1.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def health_check():
    return {
        "status": "Tsifulator.ai is running",
        "version": "0.1.0",
        "env": os.getenv("ENV", "production")
    }

# --- Routes ---
from routes import chat, auth, gmail, files, notes
app.include_router(chat.router, prefix="/chat")
app.include_router(auth.router, prefix="/auth")
app.include_router(gmail.router, prefix="/gmail")
app.include_router(files.router, prefix="/files")
app.include_router(notes.router, prefix="/notes")

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
    "notes": "Notes",
    "terminal": "Terminal",
    "safari": "Safari",
    "chrome": "Google Chrome",
    "finder": "Finder",
}

@app.post("/launch-app")
async def launch_app(request: LaunchAppRequest):
    app_key = request.app_name.lower().strip()
    app_name = APP_NAMES.get(app_key)
    if not app_name:
        return {"status": "error", "message": f"Unknown app: {request.app_name}. Available: {', '.join(set(APP_NAMES.values()))}"}
    try:
        subprocess.Popen(["open", "-a", app_name])
        return {"status": "ok", "message": f"Launching {app_name}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}
