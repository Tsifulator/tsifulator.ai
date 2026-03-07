/**
 * Shared Memory System — cross-adapter context store.
 *
 * Adapters (terminal, excel, rstudio) write and read memory entries
 * so context flows between surfaces. E.g. an Excel formula explored
 * in the Excel adapter is visible when the user switches to terminal.
 *
 * Memory entries are scoped per-user and optionally per-session.
 * Each entry has a namespace (adapter name or "global"), a key, and a value.
 * Entries expire after a configurable TTL (default 24h).
 */

import { AppDb } from "./db";

export interface MemoryEntry {
  id: string;
  userId: string;
  namespace: string;
  key: string;
  value: string;
  sessionId: string | null;
  createdAt: string;
  expiresAt: string;
}

export interface SharedMemoryOptions {
  defaultTtlMs?: number;
}

const DEFAULT_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

export class SharedMemory {
  private db: AppDb;
  private defaultTtlMs: number;

  constructor(db: AppDb, options: SharedMemoryOptions = {}) {
    this.db = db;
    this.defaultTtlMs = options.defaultTtlMs ?? DEFAULT_TTL_MS;
  }

  /** Store a memory entry visible to all adapters for this user. */
  set(userId: string, namespace: string, key: string, value: string, sessionId?: string, ttlMs?: number): MemoryEntry {
    const ttl = ttlMs ?? this.defaultTtlMs;
    const now = new Date();
    const expiresAt = new Date(now.getTime() + ttl).toISOString();
    return this.db.setMemoryEntry(userId, namespace, key, value, sessionId ?? null, expiresAt);
  }

  /** Get a specific entry by namespace + key for this user. */
  get(userId: string, namespace: string, key: string): MemoryEntry | undefined {
    return this.db.getMemoryEntry(userId, namespace, key);
  }

  /** Get all non-expired entries for a user, optionally filtered by namespace. */
  list(userId: string, namespace?: string, limit = 50): MemoryEntry[] {
    return this.db.listMemoryEntries(userId, namespace, limit);
  }

  /** Build a context string from recent memory for injection into chat. */
  buildContext(userId: string, limit = 20): string {
    const entries = this.list(userId, undefined, limit);
    if (entries.length === 0) return "";

    const lines = entries.map(
      (e) => `[${e.namespace}] ${e.key}: ${e.value}`
    );
    return `--- Shared Memory ---\n${lines.join("\n")}\n--- End Memory ---`;
  }

  /** Delete a specific entry. */
  delete(userId: string, namespace: string, key: string): boolean {
    return this.db.deleteMemoryEntry(userId, namespace, key);
  }

  /** Purge expired entries (call periodically). */
  purgeExpired(): number {
    return this.db.purgeExpiredMemoryEntries();
  }
}
