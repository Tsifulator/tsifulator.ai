import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4500 + Math.floor(Math.random() * 400);
const BASE_URL = `http://127.0.0.1:${TEST_PORT}`;
const TEST_DB_PATH = path.resolve(`./data/test-apikeys-${process.pid}-${Date.now()}.db`);

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

test("API key lifecycle: create, authenticate, list, revoke", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  // Login to get a dev token first
  const login = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: "apikey@test.ai" }),
  });
  assert.equal(login.status, 200);
  const { token: devToken } = await login.json();

  // Create an API key
  const create = await fetch(`${BASE_URL}/auth/api-keys`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${devToken}` },
    body: JSON.stringify({ name: "test-key" }),
  });
  assert.equal(create.status, 200);
  const createPayload = await create.json();
  assert.ok(createPayload.key.startsWith("tsk_"), "Key should start with tsk_");
  assert.ok(createPayload.id);
  assert.ok(createPayload.warning);

  const apiKey = createPayload.key;
  const keyId = createPayload.id;

  // Authenticate with the API key — should be able to chat
  const chat = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({ message: "hello from api key" }),
  });
  assert.equal(chat.status, 200);
  const chatPayload = await chat.json();
  assert.ok(chatPayload.sessionId, "Should get a session via API key auth");

  // List API keys
  const list = await fetch(`${BASE_URL}/auth/api-keys`, {
    headers: { authorization: `Bearer ${devToken}` },
  });
  assert.equal(list.status, 200);
  const listPayload = await list.json();
  assert.ok(listPayload.keys.length >= 1);
  assert.ok(listPayload.keys.some((k) => k.id === keyId));

  // Revoke the API key
  const revoke = await fetch(`${BASE_URL}/auth/api-keys/revoke`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${devToken}` },
    body: JSON.stringify({ keyId }),
  });
  assert.equal(revoke.status, 200);
  const revokePayload = await revoke.json();
  assert.equal(revokePayload.status, "revoked");

  // Revoked key should no longer authenticate
  const chatAfterRevoke = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({ message: "should fail" }),
  });
  assert.equal(chatAfterRevoke.status, 401, "Revoked API key should be rejected");

  // Double-revoke should 404
  const doubleRevoke = await fetch(`${BASE_URL}/auth/api-keys/revoke`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${devToken}` },
    body: JSON.stringify({ keyId }),
  });
  assert.equal(doubleRevoke.status, 404);
});
