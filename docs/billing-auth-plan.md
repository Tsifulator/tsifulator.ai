# Billing & Auth Plan

## Current State (Beta)

- **Auth**: Dev token (`POST /auth/dev-login` → `Bearer dev-<userId>`)
- **Billing**: None — free beta access
- **User model**: Email-based, single table, no password

## Launch Auth Plan

### Phase 1: Launch (Month 1-2)

**Auth provider**: Third-party (recommended: [Clerk](https://clerk.com) or [Auth0](https://auth0.com))
- JWT-based sessions replacing dev tokens
- Email + password + OAuth (Google, GitHub)
- Middleware swap: replace `requireDevAuth` with JWT verification
- User table gains `externalId` column linking to auth provider

**Migration path**:
1. Add `external_id` column to `users` table
2. Create new middleware `requireAuth` that verifies JWTs from chosen provider
3. Keep `requireDevAuth` behind `NODE_ENV !== "production"` for local dev
4. Map `externalId` → existing `userId` on first login (upsert)

### Phase 2: Post-Launch (Month 3+)

- API key auth for programmatic access
- Team/org accounts (shared sessions, shared billing)

## Launch Billing Plan

### Tier Structure

| Tier | Price | Limits |
|------|-------|--------|
| **Free** | $0/mo | 50 prompts/day, 5 sessions, community support |
| **Pro** | $19/mo | Unlimited prompts, unlimited sessions, priority support |
| **Team** | $49/mo per seat | Pro + shared sessions, audit log, SSO |

### Implementation

**Payment provider**: [Stripe](https://stripe.com)
- Stripe Checkout for subscription signup
- Stripe Customer Portal for self-service management
- Webhook listener for `customer.subscription.created/updated/deleted`

**Enforcement**:
1. Add `plan` field to users table (`free` | `pro` | `team`)
2. Add `stripe_customer_id` and `stripe_subscription_id` columns
3. Rate-limit middleware checks `plan` before processing requests
4. Free tier: count prompts in rolling 24h window via `event_log`
5. Over-limit returns `429 Too Many Requests` with upgrade URL

**DB changes**:
```sql
ALTER TABLE users ADD COLUMN plan TEXT NOT NULL DEFAULT 'free';
ALTER TABLE users ADD COLUMN stripe_customer_id TEXT;
ALTER TABLE users ADD COLUMN stripe_subscription_id TEXT;
ALTER TABLE users ADD COLUMN external_id TEXT;
CREATE INDEX idx_users_external_id ON users(external_id);
CREATE INDEX idx_users_stripe_customer_id ON users(stripe_customer_id);
```

### Launch Day Billing Sequence

1. Ship with Free tier only (rate-limited)
2. Enable Stripe checkout for Pro tier within week 1
3. Add Team tier when org features are ready (month 3+)

## API Endpoints to Add

| Method | Path | Description |
|--------|------|-------------|
| POST | `/auth/login` | Exchange auth provider token for session |
| POST | `/billing/checkout` | Create Stripe Checkout session |
| POST | `/billing/webhook` | Stripe webhook handler |
| GET | `/billing/portal` | Redirect to Stripe Customer Portal |
| GET | `/account` | Current user plan + usage |

## Open Decisions

- [ ] Choose auth provider (Clerk vs Auth0 vs custom)
- [ ] Finalize free tier daily limit (50 is placeholder)
- [ ] Decide if Pro has any soft limits or is truly unlimited
- [ ] Team billing: per-seat or flat rate?
