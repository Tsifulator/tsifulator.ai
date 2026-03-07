#!/usr/bin/env node

/**
 * Simple load test for tsifulator.ai critical flows.
 *
 * Usage: node scripts/load-test.mjs [baseUrl] [concurrency] [requests]
 *
 * Defaults: http://localhost:4000, 10 concurrent, 100 total requests
 */

const BASE_URL = process.argv[2] || "http://localhost:4000";
const CONCURRENCY = parseInt(process.argv[3] || "10", 10);
const TOTAL = parseInt(process.argv[4] || "100", 10);

async function login() {
  const res = await fetch(`${BASE_URL}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email: `loadtest-${Date.now()}@test.ai` }),
  });
  if (!res.ok) throw new Error(`Login failed: ${res.status}`);
  const data = await res.json();
  return data.token;
}

async function timedFetch(url, options) {
  const start = performance.now();
  try {
    const res = await fetch(url, options);
    const latency = performance.now() - start;
    return { status: res.status, latency, ok: res.ok };
  } catch (err) {
    const latency = performance.now() - start;
    return { status: 0, latency, ok: false, error: err.message };
  }
}

async function runBatch(name, fn, total, concurrency) {
  const results = [];
  let completed = 0;

  async function worker() {
    while (completed < total) {
      const idx = completed++;
      if (idx >= total) break;
      results.push(await fn(idx));
    }
  }

  const startTime = performance.now();
  const workers = Array.from({ length: Math.min(concurrency, total) }, () => worker());
  await Promise.all(workers);
  const totalTime = performance.now() - startTime;

  const latencies = results.map((r) => r.latency).sort((a, b) => a - b);
  const successes = results.filter((r) => r.ok).length;
  const failures = results.filter((r) => !r.ok).length;

  const p50 = latencies[Math.floor(latencies.length * 0.5)];
  const p95 = latencies[Math.floor(latencies.length * 0.95)];
  const p99 = latencies[Math.floor(latencies.length * 0.99)];
  const avg = latencies.reduce((a, b) => a + b, 0) / latencies.length;
  const rps = (results.length / totalTime) * 1000;

  console.log(`\n── ${name} ──`);
  console.log(`  Total: ${results.length} | OK: ${successes} | Fail: ${failures}`);
  console.log(`  Duration: ${(totalTime / 1000).toFixed(2)}s | RPS: ${rps.toFixed(1)}`);
  console.log(`  Latency: avg=${avg.toFixed(0)}ms p50=${p50.toFixed(0)}ms p95=${p95.toFixed(0)}ms p99=${p99.toFixed(0)}ms`);

  return { name, total: results.length, successes, failures, rps, avg, p50, p95, p99, totalTime };
}

async function main() {
  console.log(`Load test: ${BASE_URL} | concurrency=${CONCURRENCY} | requests=${TOTAL}\n`);

  // Verify server is up
  try {
    const health = await fetch(`${BASE_URL}/health`);
    if (!health.ok) throw new Error(`Health check failed: ${health.status}`);
    console.log("✅ Server is healthy");
  } catch (err) {
    console.error(`❌ Server not reachable at ${BASE_URL}: ${err.message}`);
    process.exit(1);
  }

  const token = await login();
  console.log("✅ Auth token acquired");

  // Test 1: Health endpoint (baseline)
  await runBatch("GET /health", () =>
    timedFetch(`${BASE_URL}/health`),
    TOTAL, CONCURRENCY
  );

  // Test 2: Chat endpoint (core flow)
  await runBatch("POST /chat", (i) =>
    timedFetch(`${BASE_URL}/chat`, {
      method: "POST",
      headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
      body: JSON.stringify({ message: `load test message ${i}` }),
    }),
    TOTAL, CONCURRENCY
  );

  // Test 3: Chat with command proposal
  await runBatch("POST /chat (cmd proposal)", (i) =>
    timedFetch(`${BASE_URL}/chat`, {
      method: "POST",
      headers: { "content-type": "application/json", authorization: `Bearer ${token}` },
      body: JSON.stringify({ message: `cmd: echo test-${i}` }),
    }),
    Math.min(TOTAL, 50), CONCURRENCY
  );

  // Test 4: SSE stream
  await runBatch("GET /chat/stream", (i) =>
    timedFetch(`${BASE_URL}/chat/stream?message=${encodeURIComponent(`stream test ${i}`)}`, {
      headers: { authorization: `Bearer ${token}` },
    }),
    Math.min(TOTAL, 50), Math.min(CONCURRENCY, 5)
  );

  // Test 5: Session list
  await runBatch("GET /sessions", () =>
    timedFetch(`${BASE_URL}/sessions?limit=10`, {
      headers: { authorization: `Bearer ${token}` },
    }),
    TOTAL, CONCURRENCY
  );

  console.log("\n✅ Load test complete\n");
}

main().catch((err) => {
  console.error("Load test failed:", err);
  process.exit(1);
});
