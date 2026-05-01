"""
Billing Route — Stripe subscription management for tsifl.
Handles checkout sessions, webhooks, and subscription status checks.

Flow:
  1. User clicks "Subscribe" in add-in → POST /billing/checkout → Stripe Checkout URL
  2. User pays → Stripe fires webhook → POST /billing/webhook → marks user subscribed
  3. Every /chat call → GET /billing/status/{user_id} → bypass task limits if subscribed
"""

import stripe
import os
import json
import time
import logging
from pathlib import Path
from fastapi import APIRouter, HTTPException, Request, Header
from pydantic import BaseModel
from typing import Optional

logger = logging.getLogger(__name__)

stripe.api_key = os.getenv("STRIPE_SECRET_KEY", "")
PUBLISHABLE_KEY = os.getenv("STRIPE_PUBLISHABLE_KEY", os.getenv("STRIPE_PUBLISHABLE_...", ""))

router = APIRouter()

# ── Subscription store (file-backed, survives restarts within same Railway volume) ──
_SUB_FILE = "/tmp/.tsifl_subscriptions.json"
_sub_store: dict = {}


def _load_subs():
    global _sub_store
    if Path(_SUB_FILE).exists():
        try:
            with open(_SUB_FILE) as f:
                _sub_store = json.load(f)
        except Exception:
            _sub_store = {}


def _save_subs():
    try:
        with open(_SUB_FILE, "w") as f:
            json.dump(_sub_store, f)
    except Exception:
        pass


_load_subs()

# ── Price ID cache (create product+price once, reuse) ──────────────────────────
_PRICE_ID_FILE = "/tmp/.tsifl_price_id.txt"
_cached_price_id: Optional[str] = None


def _get_or_create_price() -> str:
    global _cached_price_id

    env_price = os.getenv("STRIPE_PRICE_ID", "")
    if env_price:
        return env_price

    if Path(_PRICE_ID_FILE).exists():
        try:
            pid = Path(_PRICE_ID_FILE).read_text().strip()
            if pid:
                _cached_price_id = pid
                return pid
        except Exception:
            pass

    if _cached_price_id:
        return _cached_price_id

    if not stripe.api_key:
        raise HTTPException(status_code=500, detail="Stripe not configured")

    try:
        product = stripe.Product.create(
            name="tsifl Pro",
            description=(
                "AI-powered financial workflow layer — "
                "Excel comp builder, PowerPoint deck generator, "
                "real market data, IB formatting. Everything big IB uses, at 1% of the cost."
            ),
        )
        price = stripe.Price.create(
            product=product.id,
            unit_amount=9900,       # $99.00
            currency="usd",
            recurring={"interval": "month"},
            nickname="tsifl Pro Monthly",
        )
        _cached_price_id = price.id
        Path(_PRICE_ID_FILE).write_text(price.id)
        logger.info(f"[billing] Created Stripe product {product.id}, price {price.id}")
        return price.id
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Stripe product setup failed: {e}")


# ── Helpers ─────────────────────────────────────────────────────────────────────

def is_subscribed(user_id: str) -> bool:
    """Fast sync check — used inside chat route to bypass task limits."""
    sub = _sub_store.get(user_id)
    if not sub:
        return False
    status = sub.get("status", "inactive")
    period_end = sub.get("period_end", 0)
    now = time.time()
    if status in ("active", "trialing"):
        # period_end == 0 means we haven't got the webhook yet (just completed checkout)
        return period_end == 0 or period_end > now
    return False


# ── Request models ────────────────────────────────────────────────────────────

class CheckoutRequest(BaseModel):
    user_id: str
    email: str


# ── Routes ────────────────────────────────────────────────────────────────────

@router.get("/publishable-key")
async def get_publishable_key():
    """Return the Stripe publishable key to the frontend."""
    key = PUBLISHABLE_KEY or os.getenv("STRIPE_PUBLISHABLE_KEY", "")
    if not key:
        raise HTTPException(status_code=500, detail="Stripe publishable key not set")
    return {"publishable_key": key}


@router.post("/checkout")
async def create_checkout(req: CheckoutRequest):
    """Create a Stripe Checkout session for the $99/month subscription."""
    if not stripe.api_key:
        raise HTTPException(status_code=500, detail="Billing not configured")

    price_id = _get_or_create_price()

    try:
        session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            mode="subscription",
            line_items=[{"price": price_id, "quantity": 1}],
            customer_email=req.email,
            allow_promotion_codes=True,
            success_url=(
                "https://focused-solace-production-6839.up.railway.app"
                "/billing/success?session_id={CHECKOUT_SESSION_ID}"
            ),
            cancel_url=(
                "https://focused-solace-production-6839.up.railway.app/billing/cancel"
            ),
            metadata={"user_id": req.user_id},
            subscription_data={
                "trial_period_days": 14,
                "metadata": {"user_id": req.user_id},
            },
        )
        return {"url": session.url, "session_id": session.id}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/status/{user_id}")
async def subscription_status(user_id: str):
    """Check if a user has an active or trialing subscription."""
    subscribed = is_subscribed(user_id)
    sub = _sub_store.get(user_id, {})
    return {
        "subscribed": subscribed,
        "status": sub.get("status", "inactive"),
        "period_end": sub.get("period_end", 0),
    }


@router.post("/webhook")
async def stripe_webhook(
    request: Request,
    stripe_signature: Optional[str] = Header(None, alias="stripe-signature"),
):
    """Handle Stripe webhook events to activate/deactivate subscriptions."""
    webhook_secret = os.getenv("STRIPE_WEBHOOK_SECRET", "")
    body = await request.body()

    try:
        if webhook_secret and stripe_signature:
            event = stripe.Webhook.construct_event(body, stripe_signature, webhook_secret)
            event_type = event.type
            data_obj = event.data.object
        else:
            # Test mode without webhook secret — parse JSON directly
            raw = json.loads(body)
            event_type = raw.get("type", "")
            data_obj = raw.get("data", {}).get("object", {})
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Webhook error: {e}")

    logger.info(f"[billing] Webhook: {event_type}")

    # ── checkout.session.completed: immediately mark active ──────────────────
    if event_type == "checkout.session.completed":
        uid = _meta(data_obj, "user_id")
        sub_id = _field(data_obj, "subscription")
        if uid:
            _sub_store[uid] = {
                "stripe_subscription_id": sub_id or "",
                "status": "active",
                "period_end": 0,
            }
            _save_subs()
            logger.info(f"[billing] Activated user {uid}")

    # ── subscription created / updated ───────────────────────────────────────
    elif event_type in ("customer.subscription.created", "customer.subscription.updated"):
        uid = _meta(data_obj, "user_id")
        if uid:
            _sub_store[uid] = {
                "stripe_subscription_id": _field(data_obj, "id") or "",
                "status": _field(data_obj, "status") or "active",
                "period_end": _field(data_obj, "current_period_end") or 0,
            }
            _save_subs()

    # ── subscription deleted / cancelled ─────────────────────────────────────
    elif event_type == "customer.subscription.deleted":
        uid = _meta(data_obj, "user_id")
        if uid and uid in _sub_store:
            _sub_store[uid]["status"] = "canceled"
            _save_subs()
            logger.info(f"[billing] Cancelled user {uid}")

    return {"received": True}


@router.get("/success")
async def billing_success(session_id: str = ""):
    return {
        "status": "success",
        "message": "Subscription activated — return to tsifl and reload the add-in.",
    }


@router.get("/cancel")
async def billing_cancel():
    return {
        "status": "cancelled",
        "message": "No charge made — return to tsifl whenever you're ready.",
    }


# ── Utility ───────────────────────────────────────────────────────────────────

def _field(obj, key):
    """Get a field from either a dict or a Stripe object."""
    if isinstance(obj, dict):
        return obj.get(key)
    return getattr(obj, key, None)


def _meta(obj, key):
    """Get a metadata field from either a dict or a Stripe object."""
    meta = _field(obj, "metadata")
    if not meta:
        return None
    if isinstance(meta, dict):
        return meta.get(key)
    return getattr(meta, key, None)
