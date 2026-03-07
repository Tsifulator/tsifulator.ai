# Production User Model

## Current model (beta)

```sql
CREATE TABLE users (
  id TEXT PRIMARY KEY,        -- uid('usr_...')
  email TEXT UNIQUE NOT NULL,
  created_at TEXT NOT NULL     -- ISO 8601
);
```

TypeScript interface:
```typescript
interface AuthUser {
  id: string;
  email: string;
}
```

Auth: dev-login returns `dev-<userId>` as bearer token. No password, no hashing.

## Production model (proposed)

### Schema changes

```sql
ALTER TABLE users ADD COLUMN display_name TEXT;
ALTER TABLE users ADD COLUMN auth_provider TEXT NOT NULL DEFAULT 'dev';
ALTER TABLE users ADD COLUMN auth_provider_id TEXT;
ALTER TABLE users ADD COLUMN plan TEXT NOT NULL DEFAULT 'free';
ALTER TABLE users ADD COLUMN last_login_at TEXT;
ALTER TABLE users ADD COLUMN disabled_at TEXT;
```

### Updated interface

```typescript
interface AuthUser {
  id: string;
  email: string;
  displayName: string | null;
  authProvider: 'dev' | 'github' | 'google' | 'api-key';
  plan: 'free' | 'pro' | 'team';
  lastLoginAt: string | null;
}
```

### Token model

| Auth method | Token format | Expiry | Use case |
|---|---|---|---|
| Dev login | `dev-<userId>` | None | Development only |
| OAuth (GitHub/Google) | JWT (signed) | 24h + refresh | Web users |
| API key | `tsk_<random>` | None (revocable) | CLI / integrations |

### API key table (new)

```sql
CREATE TABLE IF NOT EXISTS api_keys (
  id TEXT PRIMARY KEY,
  user_id TEXT NOT NULL REFERENCES users(id),
  key_hash TEXT NOT NULL,          -- SHA-256 of the key
  key_prefix TEXT NOT NULL,        -- first 8 chars for identification
  name TEXT NOT NULL,              -- user-given label
  last_used_at TEXT,
  created_at TEXT NOT NULL,
  revoked_at TEXT
);
```

## Migration strategy

1. **Phase 1 (now)**: Keep dev-login, add `display_name` and `last_login_at` columns
2. **Phase 2 (pre-launch)**: Add API key support for CLI users
3. **Phase 3 (launch)**: Add OAuth provider (GitHub recommended — target audience is developers)
4. **Phase 4 (post-launch)**: Add plan/billing fields when pricing is defined

### Backward compatibility

- All migrations are additive (new columns with defaults, new tables)
- Dev-login continues to work alongside new auth methods
- Existing sessions and data remain valid
- Token validation checks `authProvider` to route to correct handler

## Rate limiting by plan

| Plan | Prompts/day | Stream requests/day |
|---|---|---|
| free | 50 | 50 |
| pro | 500 | 500 |
| team | 2000 | 2000 |

Current rate limiter already supports `maxPrompts` config — extend to read from user's plan field.
