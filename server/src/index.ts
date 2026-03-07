import { exec } from "node:child_process";
import path from "node:path";
import { promisify } from "node:util";
import cors from "@fastify/cors";
import fastifyStatic from "@fastify/static";
import Fastify from "fastify";
import { z } from "zod";
import { requireDevAuth } from "./auth";
import { listAdapters } from "./adapters/registry";
import { handleChat } from "./chat-engine";
import { getConfig } from "./config";
import { AppDb } from "./db";
import { requireRateLimit } from "./rate-limit";
import { boundOutput, classifyRisk, redactSecrets } from "./risk";
import { SharedMemory } from "./shared-memory";
import { AuthUser } from "./types";

const execAsync = promisify(exec);

const config = getConfig();
const db = new AppDb(config.DB_PATH);
const memory = new SharedMemory(db);
const app = Fastify({ logger: { level: config.LOG_LEVEL } });

// CORS — disabled when CORS_ORIGIN is empty
if (config.CORS_ORIGIN) {
  app.register(cors, {
    origin: config.CORS_ORIGIN === "*" ? true : config.CORS_ORIGIN.split(",").map((s) => s.trim()),
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    exposedHeaders: ["X-RateLimit-Limit", "X-RateLimit-Remaining", "X-RateLimit-Reset"],
  });
}

// Serve web UI from server/public
app.register(fastifyStatic, {
  root: path.join(__dirname, "..", "public"),
  prefix: "/",
  decorateReply: false,
});

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
  adapter: z.string().optional(),
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

// --- Onboarding ---

app.get("/getting-started", async () => {
  return {
    welcome: "Welcome to tsifulator.ai — your unified AI sidecar.",
    quickStart: [
      {
        step: 1,
        title: "Get a token",
        method: "POST",
        endpoint: "/auth/dev-login",
        body: { email: "you@example.com" },
        note: "Returns a bearer token for all authenticated requests.",
      },
      {
        step: 2,
        title: "Send a chat message",
        method: "POST",
        endpoint: "/chat",
        headers: { Authorization: "Bearer <your-token>" },
        body: { message: "hello" },
        note: "Returns an AI response. Prefix with 'cmd:' to get a command proposal.",
      },
      {
        step: 3,
        title: "Try a command proposal",
        method: "POST",
        endpoint: "/chat",
        headers: { Authorization: "Bearer <your-token>" },
        body: { message: "cmd: echo hello world" },
        note: "Returns a proposal with risk level. Approve it via /actions/approve.",
      },
      {
        step: 4,
        title: "Approve the command",
        method: "POST",
        endpoint: "/actions/approve",
        headers: { Authorization: "Bearer <your-token>" },
        body: { proposalId: "<from-step-3>", approved: true, cwd: "." },
        note: "Executes the command and returns output. Set approved:false to reject.",
      },
      {
        step: 5,
        title: "Stream a response (SSE)",
        method: "GET",
        endpoint: "/chat/stream?message=hello",
        headers: { Authorization: "Bearer <your-token>" },
        note: "Returns Server-Sent Events: chunk → optional proposal → done.",
      },
    ],
    apiKeys: {
      title: "Create an API key for CLI/integrations",
      create: { method: "POST", endpoint: "/auth/api-keys", body: { name: "my-key" } },
      list: { method: "GET", endpoint: "/auth/api-keys" },
      revoke: { method: "POST", endpoint: "/auth/api-keys/revoke", body: { keyId: "<id>" } },
    },
    adapters: {
      available: listAdapters(),
      usage: "Pass 'adapter' field in /chat or /chat/stream to use a specific adapter (default: terminal).",
    },
    docs: {
      devRunbook: "/docs/dev-runbook.md",
      authAndSessions: "/docs/auth-and-sessions.md",
      adapterInterfaces: "/docs/adapter-interfaces.md",
      deployTarget: "/docs/deploy-target.md",
    },
    terminal: {
      title: "Terminal CLI",
      command: "npm run cli",
      kpiMode: "npm run cli -- --kpi",
    },
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

// --- API Key Management ---

app.post("/auth/api-keys", { preHandler: authPreHandler }, async (request, reply) => {
  const schema = z.object({ name: z.string().min(1).max(100) });
  const parsed = schema.safeParse(request.body);
  if (!parsed.success) {
    return reply.code(400).send({ error: "Provide a name for the API key" });
  }

  const authUser = getAuthUser(request);
  const result = db.createApiKey(authUser.id, parsed.data.name);

  return {
    id: result.id,
    key: result.key,
    keyPrefix: result.keyPrefix,
    name: parsed.data.name,
    warning: "Save this key now — it cannot be retrieved again.",
  };
});

app.get("/auth/api-keys", { preHandler: authPreHandler }, async (request) => {
  const authUser = getAuthUser(request);
  return { keys: db.listApiKeys(authUser.id) };
});

app.post("/auth/api-keys/revoke", { preHandler: authPreHandler }, async (request, reply) => {
  const schema = z.object({ keyId: z.string().min(1) });
  const parsed = schema.safeParse(request.body);
  if (!parsed.success) {
    return reply.code(400).send({ error: "Provide keyId to revoke" });
  }

  const authUser = getAuthUser(request);
  const revoked = db.revokeApiKey(authUser.id, parsed.data.keyId);

  if (!revoked) {
    return reply.code(404).send({ error: "API key not found or already revoked" });
  }

  return { status: "revoked", keyId: parsed.data.keyId };
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
    adapter: body.adapter,
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
    adapter: z.string().optional(),
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
    adapter: body.adapter,
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

  // Send SSE keepalive comments every 15s to prevent proxies/LBs from dropping the connection
  const keepaliveInterval = setInterval(() => {
    if (!closed) {
      reply.raw.write(`: keepalive\n\n`);
    }
  }, 15_000);

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

  clearInterval(keepaliveInterval);
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
    const rawMessage = String(error);
    const output = redactSecrets(boundOutput(rawMessage));

    // Classify the failure for the client
    const isTimeout = rawMessage.includes("ETIMEDOUT") || rawMessage.includes("timed out");
    const isPermission = rawMessage.includes("EPERM") || rawMessage.includes("EACCES") || rawMessage.includes("Access is denied");
    const isNotFound = rawMessage.includes("ENOENT") || rawMessage.includes("not recognized") || rawMessage.includes("not found");

    const hint = isTimeout
      ? "Command exceeded the 30-second timeout. Try a faster alternative or break it into smaller steps."
      : isPermission
        ? "Permission denied. The command may require elevated privileges or access to a restricted path."
        : isNotFound
          ? "Command or path not found. Check spelling and ensure the tool is installed."
          : "Command failed during execution. Review the output for details.";

    db.saveExecution(proposal.id, "error", output);
    db.logEvent(proposal.sessionId, "action_failed", { proposalId: proposal.id, output });
    return reply.code(500).send({ status: "error", output, hint, proposal });
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

app.get("/telemetry/recent-events", { preHandler: authPreHandler }, async (request) => {
  const query = z
    .object({
      limit: z.coerce.number().int().min(1).max(200).default(50),
    })
    .safeParse(request.query);

  const limit = query.success ? query.data.limit : 50;
  const authUser = getAuthUser(request);
  const events = db.getRecentEventsByUser(authUser.id, limit);

  return {
    generatedAt: new Date().toISOString(),
    count: events.length,
    events: events.map((e) => ({
      ...e,
      payload: (() => { try { return JSON.parse(e.payload); } catch { return e.payload; } })(),
    })),
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

// --- Shared Memory API ---

const memorySetSchema = z.object({
  namespace: z.string().min(1).max(64),
  key: z.string().min(1).max(256),
  value: z.string().min(1).max(10000),
  sessionId: z.string().optional(),
  ttlMs: z.number().int().positive().optional(),
});

app.post("/memory", { preHandler: authPreHandler }, async (request, reply) => {
  const parsed = memorySetSchema.safeParse(request.body);
  if (!parsed.success) {
    return reply.code(400).send({ error: parsed.error.flatten() });
  }
  const authUser = getAuthUser(request);
  const entry = memory.set(
    authUser.id,
    parsed.data.namespace,
    parsed.data.key,
    parsed.data.value,
    parsed.data.sessionId,
    parsed.data.ttlMs,
  );
  return entry;
});

app.get("/memory", { preHandler: authPreHandler }, async (request) => {
  const query = z.object({
    namespace: z.string().optional(),
    limit: z.coerce.number().int().min(1).max(200).default(50),
  }).safeParse(request.query);
  const authUser = getAuthUser(request);
  const params = query.success ? query.data : { limit: 50 };
  return {
    entries: memory.list(authUser.id, params.namespace, params.limit),
  };
});

app.get("/memory/context", { preHandler: authPreHandler }, async (request) => {
  const authUser = getAuthUser(request);
  return { context: memory.buildContext(authUser.id) };
});

app.delete("/memory/:namespace/:key", { preHandler: authPreHandler }, async (request, reply) => {
  const params = z.object({ namespace: z.string(), key: z.string() }).safeParse(request.params);
  if (!params.success) return reply.code(400).send({ error: "Invalid params" });
  const authUser = getAuthUser(request);
  const deleted = memory.delete(authUser.id, params.data.namespace, params.data.key);
  return { deleted };
});

// Purge expired memory entries periodically (every 30 min)
const memoryPurgeInterval = setInterval(() => {
  try { memory.purgeExpired(); } catch {}
}, 30 * 60 * 1000);
memoryPurgeInterval.unref();

app.listen({ port: config.PORT, host: config.HOST }).then(() => {
  app.log.info(`Server running on http://localhost:${config.PORT}`);
});

// Graceful shutdown — flush SQLite WAL, close connections, finish in-flight requests
function shutdown(signal: string) {
  app.log.info(`Received ${signal}, shutting down gracefully…`);
  app.close().then(() => {
    db.close();
    app.log.info("Server closed");
    process.exit(0);
  }).catch((err) => {
    app.log.error(err, "Error during shutdown");
    process.exit(1);
  });

  // Force exit after 10s if graceful shutdown stalls
  setTimeout(() => {
    app.log.error("Graceful shutdown timed out, forcing exit");
    process.exit(1);
  }, 10_000).unref();
}

process.on("SIGTERM", () => shutdown("SIGTERM"));
process.on("SIGINT", () => shutdown("SIGINT"));
