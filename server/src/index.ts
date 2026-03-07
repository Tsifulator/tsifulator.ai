import { exec } from "node:child_process";
import { promisify } from "node:util";
import Fastify from "fastify";
import { z } from "zod";
import { requireDevAuth } from "./auth";
import { listAdapters } from "./adapters/registry";
import { handleChat } from "./chat-engine";
import { getConfig } from "./config";
import { AppDb } from "./db";
import { requireRateLimit } from "./rate-limit";
import { boundOutput, classifyRisk, redactSecrets } from "./risk";
import { AuthUser } from "./types";

const execAsync = promisify(exec);

const config = getConfig();
const db = new AppDb(config.DB_PATH);
const app = Fastify({ logger: true });
const authPreHandler = requireDevAuth(db);
const rateLimitHandler = requireRateLimit(db, {
  maxPrompts: config.RATE_LIMIT_MAX_PROMPTS,
  windowMs: config.RATE_LIMIT_WINDOW_MS,
});

function normalizeCommandForPlatform(command: string): string {
  if (process.platform === "win32") {
    return `powershell -NoProfile -ExecutionPolicy Bypass -Command "${command.replace(/"/g, '\\"')}"`;
  }
  return command;
}

const chatSchema = z.object({
  sessionId: z.string().optional(),
  message: z.string().min(1),
  cwd: z.string().optional(),
  lastOutput: z.string().optional(),
});

const approveSchema = z.object({
  proposalId: z.string(),
  approved: z.boolean(),
  cwd: z.string().optional(),
});

const feedbackSchema = z.object({
  sessionId: z.string().optional(),
  text: z.string().min(1).max(2000),
});

function getAuthUser(request: unknown): AuthUser {
  return (request as { authUser: AuthUser }).authUser;
}

// Health (existing)
app.get("/health", async () => {
  return {
    status: "ok",
    service: "tsifulator.ai",
    timestamp: new Date().toISOString(),
    adapters: listAdapters(),
  };
});

// Health (added) - so /api/health works too
app.get("/api/health", async () => {
  return {
    status: "ok",
    service: "tsifulator.ai",
    timestamp: new Date().toISOString(),
    adapters: listAdapters(),
  };
});

app.post("/auth/dev-login", async (request, reply) => {
  const schema = z.object({ email: z.string().email().default("founder@tsifulator.ai") });
  const parsed = schema.safeParse(request.body ?? {});

  if (!parsed.success) {
    return reply.code(400).send({ error: "Invalid email" });
  }

  const user = db.upsertUser(parsed.data.email);
  return {
    token: `dev-${user.id}`,
    user,
  };
});

app.post("/chat", { preHandler: [authPreHandler, rateLimitHandler] }, async (request, reply) => {
  const startedAtMs = Date.now();
  const parsed = chatSchema.safeParse(request.body);
  if (!parsed.success) {
    return reply.code(400).send({ error: parsed.error.flatten() });
  }

  const body = parsed.data;
  const authUser = getAuthUser(request);

  if (body.sessionId && !db.sessionBelongsToUser(body.sessionId, authUser.id)) {
    return reply.code(403).send({ error: "Session does not belong to authenticated user" });
  }

  const response = handleChat(db, {
    userId: authUser.id,
    sessionId: body.sessionId,
    message: body.message,
    cwd: body.cwd,
    lastOutput: body.lastOutput,
  });

  db.logEvent(response.sessionId, "chat_non_stream_completed", {
    hasProposal: Boolean(response.proposal),
    latencyMs: Date.now() - startedAtMs,
  });

  return response;
});

app.get("/chat/stream", { preHandler: [authPreHandler, rateLimitHandler] }, async (request, reply) => {
  const startedAtMs = Date.now();
  const querySchema = z.object({
    sessionId: z.string().optional(),
    message: z.string().min(1),
    cwd: z.string().optional(),
    lastOutput: z.string().optional(),
  });

  const parsed = querySchema.safeParse(request.query);
  if (!parsed.success) {
    return reply.code(400).send({ error: parsed.error.flatten() });
  }

  const body = parsed.data;
  const authUser = getAuthUser(request);

  if (body.sessionId && !db.sessionBelongsToUser(body.sessionId, authUser.id)) {
    return reply.code(403).send({ error: "Session does not belong to authenticated user" });
  }

  const payload = handleChat(db, {
    userId: authUser.id,
    sessionId: body.sessionId,
    message: body.message,
    cwd: body.cwd,
    lastOutput: body.lastOutput,
  });

  db.logEvent(payload.sessionId, "chat_stream_started", {
    hasProposal: Boolean(payload.proposal),
  });

  // SSE headers
  reply.raw.writeHead(200, {
    "Content-Type": "text/event-stream; charset=utf-8",
    "Cache-Control": "no-cache, no-transform",
    Connection: "keep-alive",
    "X-Accel-Buffering": "no",
  });

  // Some runtimes buffer headers; flush if available
  // (Fastify/node may or may not define it depending on adapters)
  (reply.raw as any).flushHeaders?.();

  // If the client disconnects, stop writing
  let closed = false;
  request.raw.on("close", () => {
    closed = true;
  });

  const chunks = payload.text.match(/.{1,40}/g) ?? [payload.text];

  let firstChunkLatencyMs: number | null = null;

  for (const chunk of chunks) {
    if (closed) break;

    if (firstChunkLatencyMs === null) {
      firstChunkLatencyMs = Date.now() - startedAtMs;
    }

    reply.raw.write(`data: ${JSON.stringify({ type: "chunk", data: chunk, sessionId: payload.sessionId })}\n\n`);
    await new Promise((resolve) => setTimeout(resolve, 80));
  }

  if (!closed && payload.proposal) {
    reply.raw.write(`data: ${JSON.stringify({ type: "proposal", data: payload.proposal })}\n\n`);
  }

  if (!closed) {
    reply.raw.write(`data: ${JSON.stringify({ type: "done", sessionId: payload.sessionId })}\n\n`);
  }

  reply.raw.end();

  db.logEvent(payload.sessionId, "chat_stream_completed", {
    chunks: chunks.length,
    hasProposal: Boolean(payload.proposal),
    firstTokenLatencyMs: firstChunkLatencyMs,
    totalLatencyMs: Date.now() - startedAtMs,
    clientClosedEarly: closed,
  });
});

app.post("/actions/approve", { preHandler: authPreHandler }, async (request, reply) => {
  const parsed = approveSchema.safeParse(request.body);
  if (!parsed.success) {
    return reply.code(400).send({ error: parsed.error.flatten() });
  }

  const body = parsed.data;
  const authUser = getAuthUser(request);

  const proposal = db.getActionProposal(body.proposalId);
  if (!proposal) {
    return reply.code(404).send({ error: "Proposal not found" });
  }

  if (!db.sessionBelongsToUser(proposal.sessionId, authUser.id)) {
    return reply.code(403).send({ error: "Proposal does not belong to authenticated user" });
  }

  if (db.hasApproval(proposal.id) || db.hasExecution(proposal.id)) {
    return reply.code(409).send({ error: "Proposal already decided" });
  }

  db.saveApproval(proposal.id, body.approved);

  if (!body.approved) {
    db.logEvent(proposal.sessionId, "action_rejected", { proposalId: proposal.id });
    return { status: "rejected", proposal };
  }

  const runtimeRisk = classifyRisk(proposal.command);
  if (proposal.risk === "blocked" || runtimeRisk === "blocked") {
    const output = "Blocked by policy";
    db.saveExecution(proposal.id, "blocked", output);
    db.logEvent(proposal.sessionId, "action_blocked", { proposalId: proposal.id });
    return { status: "blocked", output, proposal };
  }

  try {
    const executionCommand = normalizeCommandForPlatform(proposal.command);
    const { stdout, stderr } = await execAsync(executionCommand, {
      cwd: body.cwd || process.cwd(),
      windowsHide: true,
      timeout: 30_000,
      maxBuffer: 1024 * 1024,
    });

    const output = redactSecrets(boundOutput(`${stdout}\n${stderr}`.trim()));
    db.saveExecution(proposal.id, "ok", output);
    db.logEvent(proposal.sessionId, "action_executed", { proposalId: proposal.id });
    return { status: "ok", output, proposal };
  } catch (error) {
    const output = redactSecrets(boundOutput(String(error)));
    db.saveExecution(proposal.id, "error", output);
    db.logEvent(proposal.sessionId, "action_failed", { proposalId: proposal.id, output });
    return reply.code(500).send({ status: "error", output, proposal });
  }
});

app.post("/feedback", { preHandler: authPreHandler }, async (request, reply) => {
  const parsed = feedbackSchema.safeParse(request.body);
  if (!parsed.success) {
    return reply.code(400).send({ error: parsed.error.flatten() });
  }

  const authUser = getAuthUser(request);
  const { sessionId: requestedSessionId, text } = parsed.data;

  if (requestedSessionId && !db.sessionBelongsToUser(requestedSessionId, authUser.id)) {
    return reply.code(403).send({ error: "Session does not belong to authenticated user" });
  }

  const targetSessionId = requestedSessionId ?? db.createSession(authUser.id).id;

  db.logEvent(targetSessionId, "user_feedback", {
    userId: authUser.id,
    text,
    source: "terminal-cli",
  });

  return {
    status: "received",
    sessionId: targetSessionId,
  };
});

app.get("/telemetry/counters", { preHandler: authPreHandler }, async (request) => {
  const authUser = getAuthUser(request);
  const counters = db.getTelemetryCounters(authUser.id);

  return {
    generatedAt: new Date().toISOString(),
    counters,
  };
});

app.get("/sessions/:id/events", { preHandler: authPreHandler }, async (request, reply) => {
  const params = z.object({ id: z.string() }).safeParse(request.params);
  if (!params.success) {
    return reply.code(400).send({ error: "Invalid session id" });
  }

  const authUser = getAuthUser(request);
  if (!db.sessionBelongsToUser(params.data.id, authUser.id)) {
    return reply.code(403).send({ error: "Session does not belong to authenticated user" });
  }

  return {
    sessionId: params.data.id,
    events: db.getEvents(params.data.id),
  };
});

app.get("/sessions", { preHandler: authPreHandler }, async (request, reply) => {
  const query = z
    .object({
      limit: z.coerce.number().int().min(1).max(100).default(20),
    })
    .safeParse(request.query);

  if (!query.success) {
    return reply.code(400).send({ error: query.error.flatten() });
  }

  const authUser = getAuthUser(request);
  return {
    sessions: db.listSessionsByUser(authUser.id, query.data.limit),
  };
});

app.get("/sessions/search", { preHandler: authPreHandler }, async (request, reply) => {
  const query = z
    .object({
      q: z.string().min(1),
      limit: z.coerce.number().int().min(1).max(100).default(20),
    })
    .safeParse(request.query);

  if (!query.success) {
    return reply.code(400).send({ error: query.error.flatten() });
  }

  const authUser = getAuthUser(request);
  return {
    sessions: db.searchSessionsByUser(authUser.id, query.data.q, query.data.limit),
  };
});

app.get("/sessions/:id/messages", { preHandler: authPreHandler }, async (request, reply) => {
  const params = z.object({ id: z.string() }).safeParse(request.params);
  if (!params.success) {
    return reply.code(400).send({ error: "Invalid session id" });
  }

  const query = z
    .object({
      limit: z.coerce.number().int().min(1).max(500).default(100),
    })
    .safeParse(request.query);

  if (!query.success) {
    return reply.code(400).send({ error: query.error.flatten() });
  }

  const authUser = getAuthUser(request);
  if (!db.sessionBelongsToUser(params.data.id, authUser.id)) {
    return reply.code(403).send({ error: "Session does not belong to authenticated user" });
  }

  return {
    sessionId: params.data.id,
    messages: db.getMessages(params.data.id, query.data.limit),
  };
});

app.listen({ port: config.PORT, host: "0.0.0.0" }).then(() => {
  app.log.info(`Server running on http://localhost:${config.PORT}`);
});
