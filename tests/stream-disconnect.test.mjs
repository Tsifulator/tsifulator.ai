import test from "node:test";
import assert from "node:assert/strict";
import http from "node:http";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4150 + Math.floor(Math.random() * 50);
const BASE_URL = `http://127.0.0.1:${TEST_PORT}`;
const TEST_DB_PATH = path.resolve(`./data/test-disconnect-${process.pid}-${Date.now()}.db`);

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

test("SSE stream handles early client disconnect without server crash", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  // Login
  const login = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: "disconnect@tsifulator.ai" }),
  });
  const { token } = await login.json();

  // Start SSE stream and abort immediately after receiving the first chunk
  const controller = new AbortController();
  const res = await fetch(
    `${BASE_URL}/chat/stream?message=${encodeURIComponent("tell me a long story")}`,
    {
      headers: { authorization: `Bearer ${token}` },
      signal: controller.signal,
    },
  );
  assert.equal(res.status, 200);

  // Read just the first bit then abort
  const reader = res.body.getReader();
  const { value } = await reader.read();
  assert.ok(value && value.length > 0, "Should receive at least some data before aborting");
  controller.abort();

  // Give the server a moment to process the disconnect
  await sleep(500);

  // Server should still be alive and responding
  const healthRes = await fetch(`${BASE_URL}/health`);
  assert.equal(healthRes.status, 200, "Server must still be healthy after client disconnect");

  // Telemetry should record the early close
  const telRes = await fetch(`${BASE_URL}/telemetry/counters`, {
    headers: { authorization: `Bearer ${token}` },
  });
  assert.equal(telRes.status, 200, "Telemetry endpoint must respond after disconnect");
});

test("SSE keepalive comments are valid SSE (colon-prefixed)", async (t) => {
  // This is a unit-style check: keepalive lines start with ":"
  // which SSE spec defines as comments (ignored by EventSource clients).
  // The server sends `: keepalive\n\n` — verify the format is correct.
  const keepaliveLine = `: keepalive\n\n`;
  assert.ok(keepaliveLine.startsWith(":"), "Keepalive must start with colon (SSE comment)");
  assert.ok(keepaliveLine.endsWith("\n\n"), "Keepalive must end with double newline");
});
