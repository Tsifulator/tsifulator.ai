#!/usr/bin/env node

/**
 * Dev bootstrap helper — sets up a new developer environment.
 *
 * Usage: node scripts/bootstrap.mjs
 *
 * What it does:
 * 1. Copies .env.example → .env (if .env doesn't exist)
 * 2. Creates data/ directory (for SQLite)
 * 3. Runs pnpm install (if node_modules missing)
 * 4. Runs npm run build (type-check)
 * 5. Prints next steps
 */

import fs from "node:fs";
import path from "node:path";
import { execSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");

function log(msg) {
  console.log(`[bootstrap] ${msg}`);
}

function logSkip(msg) {
  console.log(`[bootstrap] ⏭  ${msg} (already exists)`);
}

// 1. .env
const envPath = path.join(root, ".env");
const envExamplePath = path.join(root, ".env.example");

if (!fs.existsSync(envPath)) {
  if (fs.existsSync(envExamplePath)) {
    fs.copyFileSync(envExamplePath, envPath);
    log("✅ Created .env from .env.example");
    log("   → Edit .env to add your OPENAI_API_KEY");
  } else {
    log("⚠️  No .env.example found — create .env manually");
  }
} else {
  logSkip(".env");
}

// 2. data directory
const dataDir = path.join(root, "data");
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
  log("✅ Created data/ directory");
} else {
  logSkip("data/");
}

// 3. Install dependencies
const nodeModules = path.join(root, "node_modules");
if (!fs.existsSync(nodeModules)) {
  log("📦 Installing dependencies...");
  try {
    execSync("pnpm install", { cwd: root, stdio: "inherit" });
    log("✅ Dependencies installed");
  } catch {
    log("⚠️  pnpm install failed — try running it manually");
  }
} else {
  logSkip("node_modules/");
}

// 4. Type-check
log("🔨 Running type-check...");
try {
  execSync("npx tsc --noEmit", { cwd: root, stdio: "inherit" });
  log("✅ Type-check passed");
} catch {
  log("⚠️  Type-check failed — fix errors before proceeding");
}

// 5. Summary
console.log("\n[bootstrap] 🚀 Ready! Next steps:");
console.log("  1. Add OPENAI_API_KEY to .env");
console.log("  2. npm run dev     — start dev server");
console.log("  3. npm run cli     — start terminal client");
console.log("  4. npm test        — run tests");
console.log("  5. See docs/dev-runbook.md for full guide\n");
