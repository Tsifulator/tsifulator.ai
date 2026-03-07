import { FastifyReply, FastifyRequest } from "fastify";
import { AppDb } from "./db";
import { AuthUser } from "./types";

export interface RateLimitConfig {
  /** Max prompts per window. 0 = unlimited. */
  maxPrompts: number;
  /** Window size in milliseconds (default: 24 hours). */
  windowMs: number;
}

const DEFAULT_CONFIG: RateLimitConfig = {
  maxPrompts: 50,
  windowMs: 24 * 60 * 60 * 1000,
};

export function requireRateLimit(db: AppDb, config: Partial<RateLimitConfig> = {}) {
  const { maxPrompts, windowMs } = { ...DEFAULT_CONFIG, ...config };

  return async function rateLimitPreHandler(request: FastifyRequest, reply: FastifyReply) {
    if (maxPrompts <= 0) return; // unlimited

    const authUser = (request as FastifyRequest & { authUser: AuthUser }).authUser;
    if (!authUser) return; // auth middleware hasn't run yet — skip

    const since = new Date(Date.now() - windowMs).toISOString();
    const count = db.countUserPromptsSince(authUser.id, since);

    // Set rate-limit headers regardless
    reply.header("X-RateLimit-Limit", maxPrompts);
    reply.header("X-RateLimit-Remaining", Math.max(0, maxPrompts - count));
    reply.header("X-RateLimit-Reset", new Date(Date.now() + windowMs).toISOString());

    if (count >= maxPrompts) {
      return reply.code(429).send({
        error: "Rate limit exceeded",
        limit: maxPrompts,
        windowMs,
        retryAfter: new Date(Date.now() + windowMs).toISOString(),
      });
    }
  };
}
