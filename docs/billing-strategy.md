# Billing Strategy

## Model: Usage-based freemium

### Tiers

| Plan | Price | Prompts/day | Streams/day | API keys | Support |
|---|---|---|---|---|---|
| **Free** | $0 | 50 | 50 | 1 | Community |
| **Pro** | $19/month | 500 | 500 | 10 | Email |
| **Team** | $49/month per seat | 2000 | 2000 | 50 | Priority |

### What counts as usage

- **Prompt**: Any `POST /chat` or `GET /chat/stream` request
- Rate limiting already exists (`RATE_LIMIT_MAX_PROMPTS` config)
- Extend rate limiter to read from user's `plan` field (see `docs/user-model-production.md`)

### Revenue drivers

1. **Pro upgrade triggers**: Hitting the 50/day free limit, needing multiple API keys, wanting faster support
2. **Team upgrade triggers**: Multiple users, shared sessions, admin controls

## Implementation plan

### Phase 1 — Soft limits (launch)

- Free tier enforced via existing rate limiter
- No payment processing yet
- "Upgrade" CTA shown when rate limit hit
- Track upgrade interest via telemetry events

### Phase 2 — Payment integration (post-launch)

- **Provider**: Stripe (best for SaaS, API-first)
- Add `stripe_customer_id` to users table
- Webhook for subscription lifecycle (created/updated/cancelled)
- Billing portal link for self-service management

### Phase 3 — Usage metering (scale)

- Track per-user usage in `event_log` (already captured)
- Monthly usage reports
- Overage billing or hard caps (TBD based on user feedback)

## Cost structure

| Component | Cost | Notes |
|---|---|---|
| OpenAI API | ~$0.002-0.06/prompt | Varies by model |
| Railway hosting | $5-20/month | Scales with usage |
| Stripe fees | 2.9% + $0.30/txn | Standard pricing |
| Domain | ~$12/year | One-time |

### Break-even estimate

At $19/month Pro plan with ~$0.03 average API cost per prompt:
- 500 prompts/day × 30 days = 15,000 prompts/month
- API cost: ~$450/month per heavy user
- **Need to cap or meter heavy usage at Pro tier**

### Mitigation

- Set hard API cost ceiling per user per month
- Use cheaper models for simple queries, expensive models only when needed
- Cache common responses where appropriate

## Pricing principles

1. **Free tier must be genuinely useful** — 50 prompts/day is enough for casual use
2. **Pro must feel worth it** — 10x the limit, priority support
3. **No surprise bills** — hard caps, not overage charges
4. **Easy to upgrade/downgrade** — self-service via Stripe billing portal
