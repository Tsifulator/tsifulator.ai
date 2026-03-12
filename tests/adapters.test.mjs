import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4500 + Math.floor(Math.random() * 50);
const BASE_URL = `http://127.0.0.1:${TEST_PORT}`;
const TEST_DB_PATH = path.resolve(`./data/test-adapters-${process.pid}-${Date.now()}.db`);

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

test("health endpoint lists all registered adapters", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  const res = await fetch(`${BASE_URL}/health`);
  assert.equal(res.status, 200);
  const payload = await res.json();
  assert.ok(Array.isArray(payload.adapters));
  assert.ok(payload.adapters.includes("terminal"), "Should have terminal adapter");
  assert.ok(payload.adapters.includes("excel"), "Should have excel adapter");
  assert.ok(payload.adapters.includes("rstudio"), "Should have rstudio adapter");
  assert.equal(payload.adapters.length, 3, "Should have exactly 3 adapters");
});
