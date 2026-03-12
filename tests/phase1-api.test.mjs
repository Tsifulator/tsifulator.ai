import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const TEST_PORT = 4100 + Math.floor(Math.random() * 50);
const BASE_URL = `http://127.0.0.1:${TEST_PORT}`;
const TEST_DB_PATH = path.resolve(`./data/test-${process.pid}-${Date.now()}.db`);

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

function startServerProcess() {
  const child = spawn(process.execPath, ["./node_modules/tsx/dist/cli.mjs", "server/src/index.ts"], {
    stdio: "ignore",
    env: {
      ...process.env,
      PORT: String(TEST_PORT),
      DB_PATH: TEST_DB_PATH
    }
  });
  return child;
}

async function stopServerProcess(child) {
  if (child.killed) {
    return;
  }

  child.kill("SIGTERM");

  await Promise.race([
    new Promise((resolve) => child.once("exit", resolve)),
    sleep(1500)
  ]);
}

function cleanupDbFiles(basePath) {
  const candidates = [basePath, `${basePath}-wal`, `${basePath}-shm`];
  for (const filePath of candidates) {
    try {
      if (fs.existsSync(filePath)) {
        fs.rmSync(filePath, { force: true });
      }
    } catch {}
  }
}

test("phase1 API auth + chat + events + ownership", async (t) => {
  const server = startServerProcess();

  t.after(async () => {
    await stopServerProcess(server);
    cleanupDbFiles(TEST_DB_PATH);
  });

  await waitForServer();

  const loginA = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: "a@tsifulator.ai" })
  });
  assert.equal(loginA.status, 200);
  const a = await loginA.json();

  const loginB = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: "b@tsifulator.ai" })
  });
  assert.equal(loginB.status, 200);
  const b = await loginB.json();

  const chat = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${a.token}`
    },
    body: JSON.stringify({ message: "cmd: Get-ChildItem" })
  });
  assert.equal(chat.status, 200);
  const chatPayload = await chat.json();
  assert.ok(chatPayload.sessionId);
  assert.ok(chatPayload.proposal?.id);

  const sessionsA = await fetch(`${BASE_URL}/sessions?limit=10`, {
    headers: {
      authorization: `Bearer ${a.token}`
    }
  });
  assert.equal(sessionsA.status, 200);
  const sessionsAPayload = await sessionsA.json();
  assert.ok(Array.isArray(sessionsAPayload.sessions));
  assert.ok(sessionsAPayload.sessions.some((item) => item.id === chatPayload.sessionId));

  const sessionsB = await fetch(`${BASE_URL}/sessions?limit=10`, {
    headers: {
      authorization: `Bearer ${b.token}`
    }
  });
  assert.equal(sessionsB.status, 200);
  const sessionsBPayload = await sessionsB.json();
  assert.ok(Array.isArray(sessionsBPayload.sessions));
  assert.equal(sessionsBPayload.sessions.some((item) => item.id === chatPayload.sessionId), false);

  const searchA = await fetch(`${BASE_URL}/sessions/search?q=childitem&limit=10`, {
    headers: {
      authorization: `Bearer ${a.token}`
    }
  });
  assert.equal(searchA.status, 200);
  const searchAPayload = await searchA.json();
  assert.ok(Array.isArray(searchAPayload.sessions));
  assert.ok(searchAPayload.sessions.some((item) => item.sessionId === chatPayload.sessionId));

  const searchB = await fetch(`${BASE_URL}/sessions/search?q=childitem&limit=10`, {
    headers: {
      authorization: `Bearer ${b.token}`
    }
  });
  assert.equal(searchB.status, 200);
  const searchBPayload = await searchB.json();
  assert.ok(Array.isArray(searchBPayload.sessions));
  assert.equal(searchBPayload.sessions.some((item) => item.sessionId === chatPayload.sessionId), false);

  const historyForeign = await fetch(`${BASE_URL}/sessions/${chatPayload.sessionId}/messages`, {
    headers: {
      authorization: `Bearer ${b.token}`
    }
  });
  assert.equal(historyForeign.status, 403);

  const historyOwn = await fetch(`${BASE_URL}/sessions/${chatPayload.sessionId}/messages?limit=50`, {
    headers: {
      authorization: `Bearer ${a.token}`
    }
  });
  assert.equal(historyOwn.status, 200);
  const historyOwnPayload = await historyOwn.json();
  assert.ok(Array.isArray(historyOwnPayload.messages));
  assert.ok(historyOwnPayload.messages.length >= 2);

  const foreignRead = await fetch(`${BASE_URL}/sessions/${chatPayload.sessionId}/events`, {
    headers: {
      authorization: `Bearer ${b.token}`
    }
  });
  assert.equal(foreignRead.status, 403);

  const foreignApprove = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${b.token}`
    },
    body: JSON.stringify({
      proposalId: chatPayload.proposal.id,
      approved: true,
      cwd: process.cwd()
    })
  });
  assert.equal(foreignApprove.status, 403);

  const ownRead = await fetch(`${BASE_URL}/sessions/${chatPayload.sessionId}/events`, {
    headers: {
      authorization: `Bearer ${a.token}`
    }
  });
  assert.equal(ownRead.status, 200);

  const approve = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${a.token}`
    },
    body: JSON.stringify({
      proposalId: chatPayload.proposal.id,
      approved: true,
      cwd: process.cwd()
    })
  });
  assert.equal(approve.status, 200);
  const approvePayload = await approve.json();
  assert.equal(approvePayload.status, "ok");

  const secondApprove = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${a.token}`
    },
    body: JSON.stringify({
      proposalId: chatPayload.proposal.id,
      approved: true,
      cwd: process.cwd()
    })
  });
  assert.equal(secondApprove.status, 409);

  const blockedChat = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${a.token}`
    },
    body: JSON.stringify({ message: "cmd: rm -rf tmp" })
  });
  assert.equal(blockedChat.status, 200);
  const blockedPayload = await blockedChat.json();
  assert.equal(blockedPayload.proposal.risk, "blocked");

  const blockedApprove = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${a.token}`
    },
    body: JSON.stringify({
      proposalId: blockedPayload.proposal.id,
      approved: true,
      cwd: process.cwd()
    })
  });
  assert.equal(blockedApprove.status, 200);
  const blockedApprovePayload = await blockedApprove.json();
  assert.equal(blockedApprovePayload.status, "blocked");

  const blockedPsChat = await fetch(`${BASE_URL}/chat`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${a.token}`
    },
    body: JSON.stringify({ message: "cmd: Remove-Item -Recurse -Force .\\tmp" })
  });
  assert.equal(blockedPsChat.status, 200);
  const blockedPsPayload = await blockedPsChat.json();
  assert.equal(blockedPsPayload.proposal.risk, "blocked");

  const blockedPsApprove = await fetch(`${BASE_URL}/actions/approve`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${a.token}`
    },
    body: JSON.stringify({
      proposalId: blockedPsPayload.proposal.id,
      approved: true,
      cwd: process.cwd()
    })
  });
  assert.equal(blockedPsApprove.status, 200);
  const blockedPsApprovePayload = await blockedPsApprove.json();
  assert.equal(blockedPsApprovePayload.status, "blocked");
});
