import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4350 + Math.floor(Math.random() * 50);
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

async function login(email) {
  const res = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email }),
  });
  assert.equal(res.status, 200);
  return res.json();
}

test("SSE stream: early client disconnect does not crash server", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  const { token } = await login("disconnect@test.ai");

  // Start a stream request and abort it immediately after first data
  const controller = new AbortController();
  const res = await fetch(
    `${BASE_URL}/chat/stream?message=${encodeURIComponent("tell me a long story")}`,
    { headers: { authorization: `Bearer ${token}` }, signal: controller.signal }
  );
  assert.equal(res.status, 200);

  // Read just the first chunk then abort
  const reader = res.body.getReader();
  const { done } = await reader.read();
  assert.equal(done, false, "Should have received at least one chunk");
  controller.abort();

  // Give the server a moment to process the disconnect
  await sleep(500);

  // Verify server is still alive and responsive
  const health = await fetch(`${BASE_URL}/health`);
  assert.equal(health.status, 200, "Server should still be healthy after client disconnect");

  // Verify we can still do normal operations
  const chat = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ message: "still works" }),
  });
  assert.equal(chat.status, 200, "Server should handle new requests after disconnect");
});
