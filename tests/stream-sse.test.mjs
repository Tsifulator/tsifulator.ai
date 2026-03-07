import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4200 + Math.floor(Math.random() * 800);
const BASE_URL = `http://127.0.0.1:${TEST_PORT}`;
const TEST_DB_PATH = path.resolve(`./data/test-stream-${process.pid}-${Date.now()}.db`);

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForServer(timeoutMs = 20000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(`${BASE_URL}/health`);
      if (res.ok) return;
    } catch {}
    await sleep(300);
  }
  throw new Error("Server did not become ready in time");
}

function startServer() {
  return spawn(process.execPath, ["./node_modules/tsx/dist/cli.mjs", "server/src/index.ts"], {
    stdio: "ignore",
    env: { ...process.env, PORT: String(TEST_PORT), DB_PATH: TEST_DB_PATH },
  });
}

async function stopServer(child) {
  if (child.killed) return;
  child.kill("SIGTERM");
  await Promise.race([new Promise((r) => child.once("exit", r)), sleep(1500)]);
}

function cleanupDb(basePath) {
  for (const p of [basePath, `${basePath}-wal`, `${basePath}-shm`]) {
    try { if (fs.existsSync(p)) fs.rmSync(p, { force: true }); } catch {}
  }
}

/** Parse an SSE text stream into an array of parsed data objects */
function parseSSE(text) {
  return text
    .split("\n\n")
    .filter((block) => block.startsWith("data: "))
    .map((block) => {
      const jsonStr = block.replace(/^data: /, "");
      return JSON.parse(jsonStr);
    });
}

test("SSE stream contract: chunks → optional proposal → done", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  // Login
  const login = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: "stream@tsifulator.ai" }),
  });
  const { token } = await login.json();

  // --- Test 1: stream a normal message (no proposal) ---
  const res1 = await fetch(
    `${BASE_URL}/chat/stream?message=${encodeURIComponent("hello world")}`,
    { headers: { authorization: `Bearer ${token}` } }
  );
  assert.equal(res1.status, 200);
  assert.ok(res1.headers.get("content-type")?.includes("text/event-stream"), "Content-Type must be text/event-stream");
  assert.equal(res1.headers.get("cache-control"), "no-cache, no-transform");

  const body1 = await res1.text();
  const events1 = parseSSE(body1);

  assert.ok(events1.length >= 2, `Expected at least 2 events (chunk + done), got ${events1.length}`);

  // All non-done events before proposal should be chunks
  const chunks1 = events1.filter((e) => e.type === "chunk");
  assert.ok(chunks1.length >= 1, "Expected at least one chunk event");
  for (const chunk of chunks1) {
    assert.ok(typeof chunk.data === "string" && chunk.data.length > 0, "Chunk data must be a non-empty string");
    assert.ok(chunk.sessionId, "Chunk must include sessionId");
  }

  // Last event must be done
  const last1 = events1[events1.length - 1];
  assert.equal(last1.type, "done", "Last event must be type=done");
  assert.ok(last1.sessionId, "Done event must include sessionId");

  // All sessionIds must match
  const sessionIds1 = new Set(events1.filter((e) => e.sessionId).map((e) => e.sessionId));
  assert.equal(sessionIds1.size, 1, "All events must share the same sessionId");

  // --- Test 2: stream a command message (should produce proposal) ---
  const res2 = await fetch(
    `${BASE_URL}/chat/stream?message=${encodeURIComponent("cmd: echo test")}`,
    { headers: { authorization: `Bearer ${token}` } }
  );
  assert.equal(res2.status, 200);

  const body2 = await res2.text();
  const events2 = parseSSE(body2);

  const proposals = events2.filter((e) => e.type === "proposal");
  assert.equal(proposals.length, 1, "Command message should produce exactly one proposal event");
  assert.ok(proposals[0].data?.id, "Proposal must have an id");
  assert.ok(proposals[0].data?.command, "Proposal must have a command");
  assert.ok(proposals[0].data?.risk, "Proposal must have a risk level");

  const last2 = events2[events2.length - 1];
  assert.equal(last2.type, "done", "Last event must be type=done even with proposal");

  // --- Test 3: reuses session when sessionId is passed ---
  const sessionId = events1.values().next().value.sessionId;
  const res3 = await fetch(
    `${BASE_URL}/chat/stream?message=${encodeURIComponent("follow up")}&sessionId=${sessionId}`,
    { headers: { authorization: `Bearer ${token}` } }
  );
  assert.equal(res3.status, 200);
  const body3 = await res3.text();
  const events3 = parseSSE(body3);
  const sessionIds3 = new Set(events3.filter((e) => e.sessionId).map((e) => e.sessionId));
  assert.ok(sessionIds3.has(sessionId), "Should reuse the provided sessionId");

  // --- Test 4: auth required ---
  const noAuth = await fetch(`${BASE_URL}/chat/stream?message=hello`);
  assert.equal(noAuth.status, 401);

  // --- Test 5: missing message returns 400 ---
  const noMsg = await fetch(`${BASE_URL}/chat/stream`, {
    headers: { authorization: `Bearer ${token}` },
  });
  assert.equal(noMsg.status, 400);
});
