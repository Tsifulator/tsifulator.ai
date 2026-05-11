"""
/agent/ — desktop agent v2 endpoint.

Native Anthropic tool calling + threaded conversation history.

Two endpoints:
  POST /agent/turn      → start or continue a conversation; returns tool_uses to run
  POST /agent/result    → post tool_result blocks back; returns next tool_uses

Conversations are kept in-memory keyed by conversation_id. Idle ones are
garbage-collected after 30 minutes.
"""

import time
import uuid
from typing import Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field

from services.agent_v2 import (
    AGENT_TOOLS,
    MODEL_STANDARD,
    MODEL_FAST,
    MODEL_DEFAULT,
    call_agent,
    append_tool_results,
    pick_model,
)
from services.cost_caps import (
    check_budget,
    record_spend,
    get_daily_spend,
    SOFT_CAP_USD,
    HARD_CAP_USD,
    PER_TURN_CAP_USD,
)

router = APIRouter()

# In-memory conversation store: {conversation_id: {"messages": [...], "last_active": ts, "context": {...}}}
_conversations: dict[str, dict] = {}
_CONV_TTL_SECONDS = 30 * 60  # 30 min


def _gc_old_conversations():
    """Drop conversations idle for more than _CONV_TTL_SECONDS."""
    now = time.time()
    stale = [
        cid for cid, c in _conversations.items()
        if now - c.get("last_active", 0) > _CONV_TTL_SECONDS
    ]
    for cid in stale:
        _conversations.pop(cid, None)


class ImageAttachment(BaseModel):
    data: str
    media_type: str
    file_name: str = ""


class TurnRequest(BaseModel):
    """Either start a new conversation (no conversation_id) or continue one."""
    user_id: str = "anon"
    conversation_id: Optional[str] = None
    message: str
    context: dict = Field(default_factory=dict)
    images: list[ImageAttachment] = Field(default_factory=list)
    # If client doesn't pin a model, server auto-picks (Haiku for simple
    # requests, Sonnet for complex). To force a specific model, set it.
    model: Optional[str] = None


class ToolResult(BaseModel):
    tool_use_id: str
    content: str
    is_error: bool = False


class ResultRequest(BaseModel):
    conversation_id: str
    results: list[ToolResult]
    context: dict = Field(default_factory=dict)  # refreshed context for next turn
    model: Optional[str] = None  # None → reuse the conversation's pinned model


class TurnResponse(BaseModel):
    conversation_id: str
    tool_uses: list[dict]
    server_tool_uses: list[dict] = []  # native Anthropic server tools (already executed)
    text: str
    thinking: str = ""
    stop_reason: str
    usage: dict
    cost_usd: float = 0.0          # cost of THIS API call
    cost_turn_usd: float = 0.0     # running cost of this user-turn (all rounds so far)
    cost_today_usd: float = 0.0    # user's running daily total
    cost_warning: str = ""         # soft/hard cap message if any
    model: str = ""
    done: bool  # true when stop_reason != "tool_use" — no more tools to run


def _make_turn_response(cid: str, result: dict, user_id: str, turn_cost: float) -> TurnResponse:
    cost = float(result.get("cost_usd", 0.0))
    spend_info = record_spend(user_id, cost)
    # Combine daily warning with per-turn warning if the turn went over
    warning = spend_info["warning"]
    if turn_cost >= PER_TURN_CAP_USD and not warning:
        warning = (
            f"💸 This request hit the per-turn cap (${turn_cost:.3f} of "
            f"${PER_TURN_CAP_USD:.2f}). Stopped to prevent runaway spend."
        )
    return TurnResponse(
        conversation_id=cid,
        tool_uses=result["tool_uses"],
        server_tool_uses=result.get("server_tool_uses", []),
        text=result["text"],
        thinking=result.get("thinking", ""),
        stop_reason=result["stop_reason"],
        usage=result["usage"],
        cost_usd=round(cost, 6),
        cost_turn_usd=round(turn_cost, 6),
        cost_today_usd=spend_info["spent"],
        cost_warning=warning,
        model=result.get("model", ""),
        done=result["stop_reason"] != "tool_use" and len(result["tool_uses"]) == 0,
    )


@router.post("/turn", response_model=TurnResponse)
async def turn(req: TurnRequest):
    """Start a new conversation or send a fresh user message to an existing one."""
    _gc_old_conversations()

    # Budget pre-flight: refuse if user is over today's hard cap.
    allowed, deny_msg = check_budget(req.user_id)
    if not allowed:
        cid = req.conversation_id or str(uuid.uuid4())
        spend = get_daily_spend(req.user_id)
        return TurnResponse(
            conversation_id=cid,
            tool_uses=[],
            text=deny_msg,
            stop_reason="budget_exceeded",
            usage={},
            cost_usd=0.0,
            cost_today_usd=spend["spent"],
            cost_warning=deny_msg,
            done=True,
        )

    cid = req.conversation_id or str(uuid.uuid4())
    conv = _conversations.get(cid, {"messages": [], "last_active": time.time()})

    # If client passed a stale conversation_id we don't know about, treat as new
    if cid not in _conversations and req.conversation_id:
        # Start fresh under the supplied id
        conv = {"messages": [], "last_active": time.time()}

    images = [img.model_dump() for img in req.images] if req.images else None

    # Auto-pick model unless client pinned one. Pin to the conversation so
    # all rounds use the same tier (consistency + caching).
    chosen_model = req.model or pick_model(req.message, has_images=bool(images))

    result = call_agent(
        user_message=req.message,
        conversation=conv["messages"],
        context=req.context,
        images=images,
        model=chosen_model,
    )

    # New user message → reset per-turn cost meter
    turn_cost = float(result.get("cost_usd", 0.0))

    # Persist updated conversation
    _conversations[cid] = {
        "messages": result["updated_conversation"],
        "last_active": time.time(),
        "context": req.context,
        "model": chosen_model,
        "user_id": req.user_id,
        "turn_cost_usd": turn_cost,
    }

    return _make_turn_response(cid, result, req.user_id, turn_cost)


@router.post("/result", response_model=TurnResponse)
async def result(req: ResultRequest):
    """Post tool_result blocks for the previous turn's tool_uses.
    Server appends them and calls Claude for the next step.
    """
    _gc_old_conversations()

    conv = _conversations.get(req.conversation_id)
    if not conv:
        raise HTTPException(
            status_code=404,
            detail=f"Unknown conversation_id: {req.conversation_id}",
        )

    user_id = conv.get("user_id", "anon")
    prior_turn_cost = float(conv.get("turn_cost_usd", 0.0))

    # Mid-loop budget check — if a runaway loop pushed us over the daily cap, stop here
    allowed, deny_msg = check_budget(user_id)
    if not allowed:
        spend = get_daily_spend(user_id)
        return TurnResponse(
            conversation_id=req.conversation_id,
            tool_uses=[],
            text=deny_msg,
            stop_reason="budget_exceeded",
            usage={},
            cost_usd=0.0,
            cost_turn_usd=prior_turn_cost,
            cost_today_usd=spend["spent"],
            cost_warning=deny_msg,
            done=True,
        )

    # Per-turn cap — has this single user message already burned too much?
    if prior_turn_cost >= PER_TURN_CAP_USD:
        msg = (
            f"💸 This request already cost ${prior_turn_cost:.3f}, hitting the "
            f"per-turn cap of ${PER_TURN_CAP_USD:.2f}. Stopping. Try a more "
            f"specific prompt, or raise TSIFL_PER_TURN_CAP_USD."
        )
        spend = get_daily_spend(user_id)
        return TurnResponse(
            conversation_id=req.conversation_id,
            tool_uses=[],
            text=msg,
            stop_reason="turn_cap_exceeded",
            usage={},
            cost_usd=0.0,
            cost_turn_usd=prior_turn_cost,
            cost_today_usd=spend["spent"],
            cost_warning=msg,
            done=True,
        )

    # Append tool_result blocks — these form the next user turn
    results_list = [
        {"tool_use_id": r.tool_use_id, "content": r.content, "is_error": r.is_error}
        for r in req.results
    ]
    updated_messages = append_tool_results(conv["messages"], results_list)

    # Reuse the conversation's pinned model unless the client overrides
    chosen_model = req.model or conv.get("model") or MODEL_DEFAULT

    # Call Claude with the appended tool_result already as the user turn.
    # build_messages skips appending an empty user message, so passing
    # user_message="" and images=None is correct.
    result_dict = call_agent(
        user_message="",
        conversation=updated_messages,
        context=req.context or conv.get("context"),
        images=None,
        model=chosen_model,
    )

    new_turn_cost = prior_turn_cost + float(result_dict.get("cost_usd", 0.0))

    _conversations[req.conversation_id] = {
        "messages": result_dict["updated_conversation"],
        "last_active": time.time(),
        "context": req.context or conv.get("context"),
        "model": chosen_model,
        "user_id": user_id,
        "turn_cost_usd": new_turn_cost,
    }

    return _make_turn_response(req.conversation_id, result_dict, user_id, new_turn_cost)


@router.delete("/conversation/{conversation_id}")
async def reset(conversation_id: str):
    """Drop a conversation explicitly. Useful for testing."""
    _conversations.pop(conversation_id, None)
    return {"ok": True}


@router.get("/health")
async def agent_health():
    """Liveness + active conversation count."""
    _gc_old_conversations()
    return {
        "ok": True,
        "active_conversations": len(_conversations),
        "tools_loaded": len(AGENT_TOOLS),
        "soft_cap_usd": SOFT_CAP_USD,
        "hard_cap_usd": HARD_CAP_USD,
        "per_turn_cap_usd": PER_TURN_CAP_USD,
    }


@router.get("/budget/{user_id}")
async def budget(user_id: str):
    """Return today's spend + remaining budget for a user."""
    return get_daily_spend(user_id)
