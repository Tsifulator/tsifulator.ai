"""
Tsifulator.ai — Backend Brain
The central server that powers the Excel add-in, RStudio panel,
and all future integrations. All chat, memory, and actions route through here.
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
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
from routes import chat, auth, gmail
app.include_router(chat.router, prefix="/chat")
app.include_router(auth.router, prefix="/auth")
app.include_router(gmail.router, prefix="/gmail")
