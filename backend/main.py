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
NOTES_APP_PATH = Path(__file__).parent.parent / "notes-app" / "index.html"

@app.get("/notes-app")
async def serve_notes_app():
    if NOTES_APP_PATH.exists():
        return FileResponse(NOTES_APP_PATH, media_type="text/html")
    return {"error": "Notes app not found"}

# --- Launch App Endpoint (for cross-app capability) ---
import subprocess
from pydantic import BaseModel

class LaunchAppRequest(BaseModel):
    app_name: str

APP_COMMANDS = {
    "microsoft excel": "open -a 'Microsoft Excel'",
    "excel": "open -a 'Microsoft Excel'",
    "microsoft powerpoint": "open -a 'Microsoft PowerPoint'",
    "powerpoint": "open -a 'Microsoft PowerPoint'",
    "microsoft word": "open -a 'Microsoft Word'",
    "word": "open -a 'Microsoft Word'",
    "visual studio code": "open -a 'Visual Studio Code'",
    "vscode": "open -a 'Visual Studio Code'",
    "vs code": "open -a 'Visual Studio Code'",
    "notes": "open -a 'Notes'",
    "terminal": "open -a 'Terminal'",
    "safari": "open -a 'Safari'",
    "chrome": "open -a 'Google Chrome'",
    "finder": "open -a 'Finder'",
}

@app.post("/launch-app")
async def launch_app(request: LaunchAppRequest):
    app_key = request.app_name.lower().strip()
    command = APP_COMMANDS.get(app_key)
    if not command:
        return {"status": "error", "message": f"Unknown app: {request.app_name}. Available: {', '.join(set(APP_COMMANDS.values()))}"}
    try:
        subprocess.Popen(command, shell=True)
        return {"status": "ok", "message": f"Launching {request.app_name}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}
