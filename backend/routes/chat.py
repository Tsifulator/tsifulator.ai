"""
Chat Route — the main entry point for all user messages.
Receives a message from Excel or RStudio, pulls memory,
sends to Claude, saves response, returns action(s).
"""

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from services.claude import get_claude_response
from services.usage import check_and_increment_usage
from services.memory import save_message, get_recent_history, is_connected

router = APIRouter()

class ChatRequest(BaseModel):
    user_id: str
    message: str
    context: dict = {}
    session_id: str = ""

class ChatResponse(BaseModel):
    reply: str
    action: dict = {}
    actions: list = []
    tasks_remaining: int = -1
    memory_active: bool = False

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

    # 2. Pull conversation history from memory (cross-app, cross-session)
    history = await get_recent_history(request.user_id, limit=10)

    # 3. Save the user's message
    app = request.context.get("app", "excel")
    await save_message(
        user_id=request.user_id,
        role="user",
        content=request.message,
        app=app,
        session_id=request.session_id
    )

    # 4. Send to Claude with full history context
    result = await get_claude_response(
        message=request.message,
        context=request.context,
        session_id=request.session_id,
        history=history
    )

    # 5. Save Claude's reply
    await save_message(
        user_id=request.user_id,
        role="assistant",
        content=result["reply"],
        app=app,
        session_id=request.session_id
    )

    return ChatResponse(
        reply=result["reply"],
        action=result.get("action", {}),
        actions=result.get("actions", []),
        tasks_remaining=usage["remaining"],
        memory_active=is_connected()
    )
