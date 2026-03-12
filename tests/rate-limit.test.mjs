import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4450 + Math.floor(Math.random() * 50);
const BASE_URL = `http://127.0.0.1:${TEST_PORT}`;
const TEST_DB_PATH = path.resolve(`./data/test-rl-${process.pid}-${Date.now()}.db`);

function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

async function waitForServer(ms = 20000) {
  const t0 = Date.now();
  while (Date.now() - t0 < ms) {
    try { if ((await fetch(`${BASE_URL}/health`)).ok) return; } catch {}
    await sleep(300);
  }
  throw new Error("Server not ready");
}

function startServer(maxPrompts) {
  return spawn(process.execPath, ["./node_modules/tsx/dist/cli.mjs", "server/src/index.ts"], {
    stdio: "ignore",
    env: {
      ...process.env,
      PORT: String(TEST_PORT),
      DB_PATH: TEST_DB_PATH,
      RATE_LIMIT_MAX_PROMPTS: String(maxPrompts),
      RATE_LIMIT_WINDOW_MS: String(60 * 60 * 1000),
    },
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

test("rate limiter returns 429 after limit exceeded", async (t) => {
  const LIMIT = 3;
  const server = startServer(LIMIT);
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  // Login
  const { token } = await (await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: "ratelimit@tsifulator.ai" }),
  })).json();

  const headers = { "content-type": "application/json", authorization: `Bearer ${token}` };

  // Send LIMIT prompts — all should succeed
  for (let i = 0; i < LIMIT; i++) {
    const res = await fetch(`${BASE_URL}/chat`, {
      method: "POST",
      headers,
      body: JSON.stringify({ message: `prompt ${i + 1}` }),
    });
    assert.equal(res.status, 200, `Prompt ${i + 1} should succeed`);
    assert.ok(res.headers.get("x-ratelimit-limit"), "Should have rate limit header");
  }

  // Next prompt should be 429
  const blocked = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers,
    body: JSON.stringify({ message: "one too many" }),
  });
  assert.equal(blocked.status, 429, "Should be rate limited");
  const body = await blocked.json();
  assert.equal(body.error, "Rate limit exceeded");
  assert.equal(body.limit, LIMIT);

  // Stream endpoint should also be 429
  const streamBlocked = await fetch(`${BASE_URL}/chat/stream?message=blocked`, {
    headers: { authorization: `Bearer ${token}` },
  });
  assert.equal(streamBlocked.status, 429, "Stream should also be rate limited");
});
