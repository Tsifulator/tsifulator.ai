import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4300 + Math.floor(Math.random() * 700);
const BASE_URL = `http://127.0.0.1:${TEST_PORT}`;
const TEST_DB_PATH = path.resolve(`./data/test-approve-${process.pid}-${Date.now()}.db`);

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

async function chatCmd(token, command, sessionId) {
  const body = { message: `cmd: ${command}` };
  if (sessionId) body.sessionId = sessionId;
  const res = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify(body),
  });
  assert.equal(res.status, 200);
  return res.json();
}

test("approve executes safe command and returns output", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  const { token } = await login("approve@test.ai");
  const chat = await chatCmd(token, "echo hello-approve-test");
  assert.ok(chat.proposal?.id);
  assert.equal(chat.proposal.risk, "safe");

  const res = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ proposalId: chat.proposal.id, approved: true, cwd: process.cwd() }),
  });
  assert.equal(res.status, 200);
  const payload = await res.json();
  assert.equal(payload.status, "ok");
  assert.ok(payload.output.includes("hello-approve-test"), `Output should contain echo text, got: ${payload.output}`);
});

test("reject records rejection and prevents later approval", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  const { token } = await login("reject@test.ai");
  const chat = await chatCmd(token, "echo should-not-run");
  assert.ok(chat.proposal?.id);

  // Reject the proposal
  const reject = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ proposalId: chat.proposal.id, approved: false }),
  });
  assert.equal(reject.status, 200);
  const rejectPayload = await reject.json();
  assert.equal(rejectPayload.status, "rejected");

  // Attempting to approve after rejection should 409
  const retry = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ proposalId: chat.proposal.id, approved: true, cwd: process.cwd() }),
  });
  assert.equal(retry.status, 409, "Should not allow approval after rejection");
});

test("blocked command cannot execute even when approved", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  const { token } = await login("blocked@test.ai");

  // rm -rf is blocked
  const chat = await chatCmd(token, "rm -rf /tmp/nope");
  assert.equal(chat.proposal.risk, "blocked");

  const res = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ proposalId: chat.proposal.id, approved: true, cwd: process.cwd() }),
  });
  assert.equal(res.status, 200);
  const payload = await res.json();
  assert.equal(payload.status, "blocked");
  assert.ok(payload.output.includes("Blocked"), "Should contain blocked message");
});

test("session messages persist across multiple chat requests", async (t) => {
  const server = startServer();
  t.after(async () => { await stopServer(server); cleanupDb(TEST_DB_PATH); });
  await waitForServer();

  const { token } = await login("persist@test.ai");

  // First message creates a session
  const chat1 = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ message: "first message" }),
  });
  assert.equal(chat1.status, 200);
  const payload1 = await chat1.json();
  const sessionId = payload1.sessionId;
  assert.ok(sessionId);

  // Second message to same session
  const chat2 = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
    body: JSON.stringify({ message: "second message", sessionId }),
  });
  assert.equal(chat2.status, 200);
  const payload2 = await chat2.json();
  assert.equal(payload2.sessionId, sessionId, "Should reuse same session");

  // Fetch history — should have messages from both requests
  const history = await fetch(`${BASE_URL}/sessions/${sessionId}/messages?limit=50`, {
    headers: { authorization: `Bearer ${token}` },
  });
  assert.equal(history.status, 200);
  const historyPayload = await history.json();
  assert.ok(historyPayload.messages.length >= 4, `Expected at least 4 messages (2 user + 2 assistant), got ${historyPayload.messages.length}`);

  const userMessages = historyPayload.messages.filter((m) => m.role === "user");
  assert.ok(userMessages.some((m) => m.content.includes("first message")));
  assert.ok(userMessages.some((m) => m.content.includes("second message")));
});
