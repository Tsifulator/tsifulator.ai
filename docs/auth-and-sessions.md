# Authentication & Sessions

## Overview

Tsifulator.ai uses a **dev-login** authentication system suitable for beta/development. Each user gets a bearer token that authenticates all API requests. Sessions are isolated per user — no user can access another user's sessions, messages, or proposals.

## Authentication flow

### 1. Get a token

```
POST /auth/dev-login
Content-Type: application/json

{ "email": "user@example.com" }
```

Response:
```json
{
  "token": "dev-abc123",
  "user": { "id": "abc123", "email": "user@example.com" }
}
```

- If the email is new, a user record is created automatically.
- If the email exists, the existing user is returned.
- The token format is `dev-<userId>`.

### 2. Use the token

Include the token in all subsequent requests:

```
Authorization: Bearer dev-abc123
```

### 3. Error responses

| Status | Error | When |
|---|---|---|
| 401 | "Authentication required" | No `Authorization` header |
| 401 | "Invalid token format" | Token doesn't start with `dev-` |
| 401 | "Token references unknown user" | User ID in token not found (stale token) |
| 403 | "Session does not belong to authenticated user" | Accessing another user's session |
| 403 | "Proposal does not belong to authenticated user" | Approving another user's proposal |

## Sessions

### How sessions work

- A **session** is created automatically when a user sends their first chat message.
- Subsequent messages can include `sessionId` to continue the same conversation.
- Sessions store: messages (user + assistant), action proposals, approvals, executions, and events.

### Session isolation

- Users can only see/access their own sessions.
- All session endpoints enforce ownership checks:
  - `GET /sessions` — lists only the authenticated user's sessions
  - `GET /sessions/search` — searches only within the authenticated user's sessions
  - `GET /sessions/:id/messages` — returns 403 if session belongs to another user
  - `GET /sessions/:id/events` — returns 403 if session belongs to another user
  - `POST /actions/approve` — returns 403 if the proposal's session belongs to another user

### Session data model

```
Session
├── id (unique)
├── userId (owner)
├── createdAt
├── Messages[] (role: user | assistant)
├── ActionProposals[] (command, risk level)
│   ├── Approval (approved: true/false)
│   └── Execution (status: ok | error | blocked, output)
└── Events[] (telemetry log)
```

## Production auth roadmap

The current `dev-login` system is intentionally simple for beta. For production:

1. **Replace dev-login** with a real auth provider (e.g., OAuth2, magic links, or API keys)
2. **Hash/sign tokens** instead of using plaintext user IDs
3. **Add token expiry** and refresh flow
4. **Add rate limiting per user** (already partially implemented via rate-limit middleware)

The session model, isolation, and authorization checks are production-ready — only the token issuance mechanism needs upgrading.
