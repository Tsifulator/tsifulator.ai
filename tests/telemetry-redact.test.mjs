import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4300 + Math.floor(Math.random() * 50);
const BASE_URL = `http://127.0.0.1:${TEST_PORT}`;
const TEST_DB_PATH = path.resolve(`./data/test-tel-${process.pid}-${Date.now()}.db`);

function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

async function waitForServer(ms = 20000) {
  const t0 = Date.now();
  while (Date.now() - t0 < ms) {
    try { if ((await fetch(`${BASE_URL}/health`)).ok) return; } catch {}
    await sleep(300);
  }
  throw new Error("Server not ready");
}

function startServer() {
  return spawn(process.execPath, ["./node_modules/tsx/dist/cli.mjs", "server/src/index.ts"], {
    stdio: "ignore",
    env: { ...process.env, PORT: String(TEST_PORT), DB_PATH: TEST_DB_PATH },
  });
}

async function stopServer(c) {
  if (c.killed) return;
  c.kill("SIGTERM");
  await Promise.race([new Promise((r) => c.once("exit", r)), sleep(1500)]);
}

function cleanupDb(p) {
  for (const f of [p, `${p}-wal`, `${p}-shm`]) {
    try { if (fs.existsSync(f)) fs.rmSync(f, { force: true }); } catch {}
  }
}

test("telemetry counters reflect activity", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  const login = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: "tel@tsifulator.ai" }),
  });
  const { token } = await login.json();

  // Baseline counters
  const c0 = await (await fetch(`${BASE_URL}/telemetry/counters`, {
    headers: { authorization: `Bearer ${token}` },
  })).json();
  assert.equal(c0.counters.promptsSent, 0);

  // Send a chat
  await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ message: "hello" }),
  });

  // Send a stream
  await (await fetch(`${BASE_URL}/chat/stream?message=hi`, {
    headers: { authorization: `Bearer ${token}` },
  })).text();

  // Check counters increased
  const c1 = await (await fetch(`${BASE_URL}/telemetry/counters`, {
    headers: { authorization: `Bearer ${token}` },
  })).json();

  assert.equal(c1.counters.promptsSent, 2, "Should have 2 prompts sent");
  assert.equal(c1.counters.promptsSent24h, 2);
  assert.equal(c1.counters.streamRequests, 1, "Should have 1 stream request");
  assert.equal(c1.counters.streamCompletions, 1, "Should have 1 stream completion");
  assert.equal(c1.counters.streamSuccessRate, 1);
  assert.ok(typeof c1.counters.medianChatLatencyMs === "number");
  assert.ok(typeof c1.generatedAt === "string");
});

test("secret redaction in command output", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  const login = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: "redact@tsifulator.ai" }),
  });
  const { token } = await login.json();

  // Send a command that echoes something that looks like a secret
  const chat = await (await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ message: "cmd: echo api_key=sk-secret123 token=abc password=hunter2" }),
  })).json();

  // Approve and execute
  const result = await (await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ proposalId: chat.proposal.id, approved: true }),
  })).json();

  assert.equal(result.status, "ok");
  // Output should have secrets redacted
  assert.ok(!result.output.includes("sk-secret123"), "API key should be redacted");
  assert.ok(!result.output.includes("hunter2"), "Password should be redacted");
  assert.ok(result.output.includes("[REDACTED]"), "Should contain [REDACTED] markers");
});
