import { createHash, randomBytes } from "node:crypto";
import Database from "better-sqlite3";

export type Role = "user" | "assistant";

export interface SessionRecord {
  id: string;
  userId: string;
}

export interface UserRecord {
  id: string;
  email: string;
}

export interface MessageRecord {
  id: string;
  sessionId: string;
  role: Role;
  content: string;
}

export interface ActionProposalRecord {
  id: string;
  sessionId: string;
  command: string;
  risk: "safe" | "confirm" | "blocked";
}

export interface SessionSummaryRecord {
  id: string;
  createdAt: string;
  lastActivityAt: string;
  messageCount: number;
}

export interface SessionMessageRecord {
  id: string;
  role: Role;
  content: string;
  createdAt: string;
}

export interface SessionSearchRecord {
  sessionId: string;
  createdAt: string;
  lastActivityAt: string;
  messageCount: number;
  snippet: string;
}

export interface TelemetryCountersRecord {
  newBetaUsers7d: number;
  dailyActiveUsers: number;
  promptsSent: number;
  promptsSent24h: number;
  applyActionsProposed: number;
  applyActionsConfirmed: number;
  blockedCommandAttempts: number;
  streamRequests: number;
  streamCompletions: number;
  streamSuccessRate: number;
  medianChatLatencyMs: number;
  medianStreamFirstTokenLatencyMs: number;
}

function uid(prefix: string): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

function median(values: number[]): number {
  if (values.length === 0) {
    return 0;
  }

  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);

  if (sorted.length % 2 === 1) {
    return sorted[middle];
  }

  return (sorted[middle - 1] + sorted[middle]) / 2;
}

function parseLatencyValue(payloadText: string, key: "latencyMs" | "firstTokenLatencyMs"): number | null {
  try {
    const parsed = JSON.parse(payloadText) as { latencyMs?: unknown; firstTokenLatencyMs?: unknown };
    const rawValue = parsed[key];

    if (typeof rawValue !== "number" || !Number.isFinite(rawValue) || rawValue < 0) {
      return null;
    }

    return rawValue;
  } catch {
    return null;
  }
}

export class AppDb {
  private db: Database.Database;

  constructor(dbPath: string) {
    this.db = new Database(dbPath);
    this.db.pragma("journal_mode = WAL");
    this.migrate();
  }

  close(): void {
    this.db.close();
  }

  private migrate(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS users (
        id TEXT PRIMARY KEY,
        email TEXT UNIQUE NOT NULL,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS sessions (
        id TEXT PRIMARY KEY,
        user_id TEXT NOT NULL,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS messages (
        id TEXT PRIMARY KEY,
        session_id TEXT NOT NULL,
        role TEXT NOT NULL,
        content TEXT NOT NULL,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS action_proposals (
        id TEXT PRIMARY KEY,
        session_id TEXT NOT NULL,
        command TEXT NOT NULL,
        risk TEXT NOT NULL,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS approvals (
        id TEXT PRIMARY KEY,
        proposal_id TEXT NOT NULL,
        approved INTEGER NOT NULL,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS action_executions (
        id TEXT PRIMARY KEY,
        proposal_id TEXT NOT NULL,
        status TEXT NOT NULL,
        output TEXT NOT NULL,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS event_log (
        id TEXT PRIMARY KEY,
        session_id TEXT NOT NULL,
        event_type TEXT NOT NULL,
        payload TEXT NOT NULL,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS adapter_states (
        id TEXT PRIMARY KEY,
        adapter TEXT NOT NULL,
        state_json TEXT NOT NULL,
        created_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS api_keys (
        id TEXT PRIMARY KEY,
        user_id TEXT NOT NULL,
        key_hash TEXT NOT NULL,
        key_prefix TEXT NOT NULL,
        name TEXT NOT NULL,
        last_used_at TEXT,
        created_at TEXT NOT NULL,
        revoked_at TEXT
      );

      CREATE TABLE IF NOT EXISTS shared_memory (
        id TEXT PRIMARY KEY,
        user_id TEXT NOT NULL,
        namespace TEXT NOT NULL,
        key TEXT NOT NULL,
        value TEXT NOT NULL,
        session_id TEXT,
        created_at TEXT NOT NULL,
        expires_at TEXT NOT NULL
      );

      CREATE UNIQUE INDEX IF NOT EXISTS idx_shared_memory_user_ns_key
        ON shared_memory(user_id, namespace, key);

      CREATE INDEX IF NOT EXISTS idx_shared_memory_expires
        ON shared_memory(expires_at);
    `);

    this.repairApprovalsDuplicates();
    this.repairExecutionDuplicates();

    this.db.exec("CREATE UNIQUE INDEX IF NOT EXISTS idx_approvals_proposal_id ON approvals(proposal_id);");
    this.db.exec("CREATE UNIQUE INDEX IF NOT EXISTS idx_action_executions_proposal_id ON action_executions(proposal_id);");
  }

  private repairApprovalsDuplicates(): void {
    const duplicateCount = (this.db
      .prepare(
        `SELECT COUNT(*) as count
         FROM (
           SELECT proposal_id
           FROM approvals
           GROUP BY proposal_id
           HAVING COUNT(*) > 1
         )`
      )
      .get() as { count: number }).count;

    if (duplicateCount <= 0) {
      return;
    }

    const tx = this.db.transaction(() => {
      this.db.exec(`
        DROP TABLE IF EXISTS approvals_dedupe_tmp;
        CREATE TABLE approvals_dedupe_tmp (
          id TEXT PRIMARY KEY,
          proposal_id TEXT NOT NULL,
          approved INTEGER NOT NULL,
          created_at TEXT NOT NULL
        );

        INSERT INTO approvals_dedupe_tmp (id, proposal_id, approved, created_at)
        SELECT a.id, a.proposal_id, a.approved, a.created_at
        FROM approvals a
        INNER JOIN (
          SELECT proposal_id, MIN(rowid) AS keep_rowid
          FROM approvals
          GROUP BY proposal_id
        ) k ON k.keep_rowid = a.rowid;

        DELETE FROM approvals;

        INSERT INTO approvals (id, proposal_id, approved, created_at)
        SELECT id, proposal_id, approved, created_at
        FROM approvals_dedupe_tmp;

        DROP TABLE approvals_dedupe_tmp;
      `);
    });

    tx();
  }

  private repairExecutionDuplicates(): void {
    const duplicateCount = (this.db
      .prepare(
        `SELECT COUNT(*) as count
         FROM (
           SELECT proposal_id
           FROM action_executions
           GROUP BY proposal_id
           HAVING COUNT(*) > 1
         )`
      )
      .get() as { count: number }).count;

    if (duplicateCount <= 0) {
      return;
    }

    const tx = this.db.transaction(() => {
      this.db.exec(`
        DROP TABLE IF EXISTS action_executions_dedupe_tmp;
        CREATE TABLE action_executions_dedupe_tmp (
          id TEXT PRIMARY KEY,
          proposal_id TEXT NOT NULL,
          status TEXT NOT NULL,
          output TEXT NOT NULL,
          created_at TEXT NOT NULL
        );

        INSERT INTO action_executions_dedupe_tmp (id, proposal_id, status, output, created_at)
        SELECT ae.id, ae.proposal_id, ae.status, ae.output, ae.created_at
        FROM action_executions ae
        INNER JOIN (
          SELECT proposal_id, MIN(rowid) AS keep_rowid
          FROM action_executions
          GROUP BY proposal_id
        ) k ON k.keep_rowid = ae.rowid;

        DELETE FROM action_executions;

        INSERT INTO action_executions (id, proposal_id, status, output, created_at)
        SELECT id, proposal_id, status, output, created_at
        FROM action_executions_dedupe_tmp;

        DROP TABLE action_executions_dedupe_tmp;
      `);
    });

    tx();
  }

  upsertUser(email: string): { id: string; email: string } {
    const existing = this.db
      .prepare("SELECT id, email FROM users WHERE email = ?")
      .get(email) as { id: string; email: string } | undefined;

    if (existing) {
      return existing;
    }

    const id = uid("usr");
    this.db
      .prepare("INSERT INTO users (id, email, created_at) VALUES (?, ?, ?)")
      .run(id, email, new Date().toISOString());

    return { id, email };
  }

  getUserById(id: string): UserRecord | undefined {
    return this.db
      .prepare("SELECT id, email FROM users WHERE id = ?")
      .get(id) as UserRecord | undefined;
  }

  // --- API Keys ---

  createApiKey(userId: string, name: string): { id: string; key: string; keyPrefix: string } {
    const id = uid("apk");
    const rawKey = `tsk_${randomBytes(24).toString("base64url")}`;
    const keyHash = createHash("sha256").update(rawKey).digest("hex");
    const keyPrefix = rawKey.slice(0, 12);

    this.db
      .prepare("INSERT INTO api_keys (id, user_id, key_hash, key_prefix, name, created_at) VALUES (?, ?, ?, ?, ?, ?)")
      .run(id, userId, keyHash, keyPrefix, name, new Date().toISOString());

    return { id, key: rawKey, keyPrefix };
  }

  validateApiKey(rawKey: string): UserRecord | undefined {
    const keyHash = createHash("sha256").update(rawKey).digest("hex");
    const row = this.db
      .prepare(
        `SELECT ak.id as keyId, ak.user_id, u.id, u.email
         FROM api_keys ak
         INNER JOIN users u ON u.id = ak.user_id
         WHERE ak.key_hash = ? AND ak.revoked_at IS NULL`
      )
      .get(keyHash) as { keyId: string; user_id: string; id: string; email: string } | undefined;

    if (!row) {
      return undefined;
    }

    // Update last_used_at
    this.db
      .prepare("UPDATE api_keys SET last_used_at = ? WHERE id = ?")
      .run(new Date().toISOString(), row.keyId);

    return { id: row.id, email: row.email };
  }

  listApiKeys(userId: string): Array<{ id: string; keyPrefix: string; name: string; lastUsedAt: string | null; createdAt: string; revokedAt: string | null }> {
    return this.db
      .prepare(
        `SELECT id, key_prefix as keyPrefix, name, last_used_at as lastUsedAt, created_at as createdAt, revoked_at as revokedAt
         FROM api_keys
         WHERE user_id = ?
         ORDER BY created_at DESC`
      )
      .all(userId) as Array<{ id: string; keyPrefix: string; name: string; lastUsedAt: string | null; createdAt: string; revokedAt: string | null }>;
  }

  revokeApiKey(userId: string, keyId: string): boolean {
    const result = this.db
      .prepare("UPDATE api_keys SET revoked_at = ? WHERE id = ? AND user_id = ? AND revoked_at IS NULL")
      .run(new Date().toISOString(), keyId, userId);
    return result.changes > 0;
  }

  countUserPromptsSince(userId: string, since: string): number {
    return (this.db
      .prepare(
        `SELECT COUNT(*) as count
         FROM messages m
         INNER JOIN sessions s ON s.id = m.session_id
         WHERE s.user_id = ? AND m.role = 'user' AND m.created_at >= ?`
      )
      .get(userId, since) as { count: number }).count;
  }

  createSession(userId: string): SessionRecord {
    const id = uid("ses");
    this.db
      .prepare("INSERT INTO sessions (id, user_id, created_at) VALUES (?, ?, ?)")
      .run(id, userId, new Date().toISOString());
    return { id, userId };
  }

  listSessionsByUser(userId: string, limit = 20): SessionSummaryRecord[] {
    return this.db
      .prepare(
        `SELECT s.id,
                s.created_at as createdAt,
                COALESCE(MAX(m.created_at), s.created_at) as lastActivityAt,
                COUNT(m.id) as messageCount
         FROM sessions s
         LEFT JOIN messages m ON m.session_id = s.id
         WHERE s.user_id = ?
         GROUP BY s.id, s.created_at
         ORDER BY lastActivityAt DESC
         LIMIT ?`
      )
      .all(userId, limit) as SessionSummaryRecord[];
  }

  getMessages(sessionId: string, limit = 100): SessionMessageRecord[] {
    return this.db
      .prepare(
        `SELECT id,
                role,
                content,
                created_at as createdAt
         FROM messages
         WHERE session_id = ?
         ORDER BY created_at ASC
         LIMIT ?`
      )
      .all(sessionId, limit) as SessionMessageRecord[];
  }

  searchSessionsByUser(userId: string, query: string, limit = 20): SessionSearchRecord[] {
    const needle = `%${query}%`;

    return this.db
      .prepare(
        `SELECT s.id as sessionId,
                s.created_at as createdAt,
                COALESCE(MAX(m.created_at), s.created_at) as lastActivityAt,
                COUNT(m.id) as messageCount,
                COALESCE(
                  MAX(CASE WHEN LOWER(m.content) LIKE LOWER(?) THEN SUBSTR(m.content, 1, 160) END),
                  ''
                ) as snippet
         FROM sessions s
         LEFT JOIN messages m ON m.session_id = s.id
         WHERE s.user_id = ?
           AND EXISTS (
             SELECT 1
             FROM messages m2
             WHERE m2.session_id = s.id
               AND LOWER(m2.content) LIKE LOWER(?)
           )
         GROUP BY s.id, s.created_at
         ORDER BY lastActivityAt DESC
         LIMIT ?`
      )
      .all(needle, userId, needle, limit) as SessionSearchRecord[];
  }

  sessionBelongsToUser(sessionId: string, userId: string): boolean {
    const row = this.db
      .prepare("SELECT id FROM sessions WHERE id = ? AND user_id = ?")
      .get(sessionId, userId) as { id: string } | undefined;

    return Boolean(row);
  }

  saveMessage(sessionId: string, role: Role, content: string): MessageRecord {
    const id = uid("msg");
    this.db
      .prepare("INSERT INTO messages (id, session_id, role, content, created_at) VALUES (?, ?, ?, ?, ?)")
      .run(id, sessionId, role, content, new Date().toISOString());
    return { id, sessionId, role, content };
  }

  saveActionProposal(sessionId: string, command: string, risk: "safe" | "confirm" | "blocked"): ActionProposalRecord {
    const id = uid("act");
    this.db
      .prepare("INSERT INTO action_proposals (id, session_id, command, risk, created_at) VALUES (?, ?, ?, ?, ?)")
      .run(id, sessionId, command, risk, new Date().toISOString());
    return { id, sessionId, command, risk };
  }

  getActionProposal(id: string): ActionProposalRecord | undefined {
    return this.db
      .prepare("SELECT id, session_id as sessionId, command, risk FROM action_proposals WHERE id = ?")
      .get(id) as ActionProposalRecord | undefined;
  }

  actionProposalBelongsToUser(proposalId: string, userId: string): boolean {
    const row = this.db
      .prepare(
        `SELECT ap.id
         FROM action_proposals ap
         INNER JOIN sessions s ON s.id = ap.session_id
         WHERE ap.id = ? AND s.user_id = ?`
      )
      .get(proposalId, userId) as { id: string } | undefined;

    return Boolean(row);
  }

  hasApproval(proposalId: string): boolean {
    const row = this.db
      .prepare("SELECT id FROM approvals WHERE proposal_id = ?")
      .get(proposalId) as { id: string } | undefined;

    return Boolean(row);
  }

  hasExecution(proposalId: string): boolean {
    const row = this.db
      .prepare("SELECT id FROM action_executions WHERE proposal_id = ?")
      .get(proposalId) as { id: string } | undefined;

    return Boolean(row);
  }

  saveApproval(proposalId: string, approved: boolean): void {
    this.db
      .prepare("INSERT INTO approvals (id, proposal_id, approved, created_at) VALUES (?, ?, ?, ?)")
      .run(uid("apr"), proposalId, approved ? 1 : 0, new Date().toISOString());
  }

  saveExecution(proposalId: string, status: "ok" | "blocked" | "error", output: string): void {
    this.db
      .prepare("INSERT INTO action_executions (id, proposal_id, status, output, created_at) VALUES (?, ?, ?, ?, ?)")
      .run(uid("exe"), proposalId, status, output, new Date().toISOString());
  }

  logEvent(sessionId: string, eventType: string, payload: unknown): void {
    this.db
      .prepare("INSERT INTO event_log (id, session_id, event_type, payload, created_at) VALUES (?, ?, ?, ?, ?)")
      .run(uid("evt"), sessionId, eventType, JSON.stringify(payload), new Date().toISOString());
  }

  getEvents(sessionId: string): Array<{ id: string; type: string; payload: string; createdAt: string }> {
    return this.db
      .prepare("SELECT id, event_type as type, payload, created_at as createdAt FROM event_log WHERE session_id = ? ORDER BY created_at ASC")
      .all(sessionId) as Array<{ id: string; type: string; payload: string; createdAt: string }>;
  }

  getRecentEventsByUser(userId: string, limit: number): Array<{ id: string; sessionId: string; type: string; payload: string; createdAt: string }> {
    return this.db
      .prepare(
        `SELECT e.id, e.session_id as sessionId, e.event_type as type, e.payload, e.created_at as createdAt
         FROM event_log e
         INNER JOIN sessions s ON s.id = e.session_id
         WHERE s.user_id = ?
         ORDER BY e.created_at DESC
         LIMIT ?`
      )
      .all(userId, limit) as Array<{ id: string; sessionId: string; type: string; payload: string; createdAt: string }>;
  }

  saveAdapterState(adapter: string, stateJson: string): void {
    this.db
      .prepare("INSERT INTO adapter_states (id, adapter, state_json, created_at) VALUES (?, ?, ?, ?)")
      .run(uid("ads"), adapter, stateJson, new Date().toISOString());
  }

  // --- Shared Memory ---

  setMemoryEntry(
    userId: string,
    namespace: string,
    key: string,
    value: string,
    sessionId: string | null,
    expiresAt: string
  ): { id: string; userId: string; namespace: string; key: string; value: string; sessionId: string | null; createdAt: string; expiresAt: string } {
    const id = uid("mem");
    const now = new Date().toISOString();
    this.db
      .prepare(
        `INSERT INTO shared_memory (id, user_id, namespace, key, value, session_id, created_at, expires_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(user_id, namespace, key) DO UPDATE SET
           value = excluded.value,
           session_id = excluded.session_id,
           created_at = excluded.created_at,
           expires_at = excluded.expires_at`
      )
      .run(id, userId, namespace, key, value, sessionId, now, expiresAt);

    // Return the actual row (may have kept old id on conflict)
    return this.getMemoryEntry(userId, namespace, key)!;
  }

  getMemoryEntry(
    userId: string,
    namespace: string,
    key: string
  ): { id: string; userId: string; namespace: string; key: string; value: string; sessionId: string | null; createdAt: string; expiresAt: string } | undefined {
    const now = new Date().toISOString();
    return this.db
      .prepare(
        `SELECT id, user_id as userId, namespace, key, value, session_id as sessionId, created_at as createdAt, expires_at as expiresAt
         FROM shared_memory
         WHERE user_id = ? AND namespace = ? AND key = ? AND expires_at > ?`
      )
      .get(userId, namespace, key, now) as any;
  }

  listMemoryEntries(
    userId: string,
    namespace?: string,
    limit = 50
  ): Array<{ id: string; userId: string; namespace: string; key: string; value: string; sessionId: string | null; createdAt: string; expiresAt: string }> {
    const now = new Date().toISOString();
    if (namespace) {
      return this.db
        .prepare(
          `SELECT id, user_id as userId, namespace, key, value, session_id as sessionId, created_at as createdAt, expires_at as expiresAt
           FROM shared_memory
           WHERE user_id = ? AND namespace = ? AND expires_at > ?
           ORDER BY created_at DESC
           LIMIT ?`
        )
        .all(userId, namespace, now, limit) as any;
    }
    return this.db
      .prepare(
        `SELECT id, user_id as userId, namespace, key, value, session_id as sessionId, created_at as createdAt, expires_at as expiresAt
         FROM shared_memory
         WHERE user_id = ? AND expires_at > ?
         ORDER BY created_at DESC
         LIMIT ?`
      )
      .all(userId, now, limit) as any;
  }

  deleteMemoryEntry(userId: string, namespace: string, key: string): boolean {
    const result = this.db
      .prepare("DELETE FROM shared_memory WHERE user_id = ? AND namespace = ? AND key = ?")
      .run(userId, namespace, key);
    return result.changes > 0;
  }

  purgeExpiredMemoryEntries(): number {
    const now = new Date().toISOString();
    const result = this.db
      .prepare("DELETE FROM shared_memory WHERE expires_at <= ?")
      .run(now);
    return result.changes;
  }

  getTelemetryCounters(userId: string): TelemetryCountersRecord {
    const now = new Date();
    const dayAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000).toISOString();
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000).toISOString();

    const newBetaUsers7d =
      (this.db
        .prepare("SELECT COUNT(*) as count FROM users WHERE created_at >= ?")
        .get(weekAgo) as { count: number }).count ?? 0;

    const dailyActiveUsers =
      (this.db
        .prepare(
          `SELECT COUNT(DISTINCT s.user_id) as count
           FROM messages m
           INNER JOIN sessions s ON s.id = m.session_id
           WHERE m.role = 'user' AND m.created_at >= ?`
        )
        .get(dayAgo) as { count: number }).count ?? 0;

    const promptsSent =
      (this.db
        .prepare(
          `SELECT COUNT(*) as count
           FROM messages m
           INNER JOIN sessions s ON s.id = m.session_id
           WHERE s.user_id = ? AND m.role = 'user'`
        )
        .get(userId) as { count: number }).count ?? 0;

    const promptsSent24h =
      (this.db
        .prepare(
          `SELECT COUNT(*) as count
           FROM messages m
           INNER JOIN sessions s ON s.id = m.session_id
           WHERE s.user_id = ? AND m.role = 'user' AND m.created_at >= ?`
        )
        .get(userId, dayAgo) as { count: number }).count ?? 0;

    const applyActionsProposed =
      (this.db
        .prepare(
          `SELECT COUNT(*) as count
           FROM action_proposals ap
           INNER JOIN sessions s ON s.id = ap.session_id
           WHERE s.user_id = ?`
        )
        .get(userId) as { count: number }).count ?? 0;

    const applyActionsConfirmed =
      (this.db
        .prepare(
          `SELECT COUNT(*) as count
           FROM approvals a
           INNER JOIN action_proposals ap ON ap.id = a.proposal_id
           INNER JOIN sessions s ON s.id = ap.session_id
           WHERE s.user_id = ? AND a.approved = 1`
        )
        .get(userId) as { count: number }).count ?? 0;

    const blockedCommandAttempts =
      (this.db
        .prepare(
          `SELECT COUNT(*) as count
           FROM action_executions ae
           INNER JOIN action_proposals ap ON ap.id = ae.proposal_id
           INNER JOIN sessions s ON s.id = ap.session_id
           WHERE s.user_id = ? AND ae.status = 'blocked'`
        )
        .get(userId) as { count: number }).count ?? 0;

    const streamRequests =
      (this.db
        .prepare(
          `SELECT COUNT(*) as count
           FROM event_log e
           INNER JOIN sessions s ON s.id = e.session_id
           WHERE s.user_id = ? AND e.event_type = 'chat_stream_started'`
        )
        .get(userId) as { count: number }).count ?? 0;

    const streamCompletions =
      (this.db
        .prepare(
          `SELECT COUNT(*) as count
           FROM event_log e
           INNER JOIN sessions s ON s.id = e.session_id
           WHERE s.user_id = ? AND e.event_type = 'chat_stream_completed'`
        )
        .get(userId) as { count: number }).count ?? 0;

    const streamSuccessRate = streamRequests > 0 ? streamCompletions / streamRequests : 1;

    const chatLatencyRows = this.db
      .prepare(
        `SELECT e.payload as payload
         FROM event_log e
         INNER JOIN sessions s ON s.id = e.session_id
         WHERE s.user_id = ? AND e.event_type = 'chat_non_stream_completed'`
      )
      .all(userId) as Array<{ payload: string }>;

    const medianChatLatencyMs = median(
      chatLatencyRows
        .map((row) => parseLatencyValue(row.payload, "latencyMs"))
        .filter((value): value is number => value !== null)
    );

    const streamLatencyRows = this.db
      .prepare(
        `SELECT e.payload as payload
         FROM event_log e
         INNER JOIN sessions s ON s.id = e.session_id
         WHERE s.user_id = ? AND e.event_type = 'chat_stream_completed'`
      )
      .all(userId) as Array<{ payload: string }>;

    const medianStreamFirstTokenLatencyMs = median(
      streamLatencyRows
        .map((row) => parseLatencyValue(row.payload, "firstTokenLatencyMs"))
        .filter((value): value is number => value !== null)
    );

    return {
      newBetaUsers7d,
      dailyActiveUsers,
      promptsSent,
      promptsSent24h,
      applyActionsProposed,
      applyActionsConfirmed,
      blockedCommandAttempts,
      streamRequests,
      streamCompletions,
      streamSuccessRate,
      medianChatLatencyMs,
      medianStreamFirstTokenLatencyMs
    };
  }
}
