import readline from "node:readline/promises";
import fs from "node:fs";
import path from "node:path";
import { gunzipSync, gzipSync } from "node:zlib";
import { stdin as input, stdout as output } from "node:process";

interface DevLoginResponse {
  token: string;
  user: { id: string; email: string };
}

interface ChatChunkEvent {
  type: "chunk" | "proposal" | "done";
  data?: unknown;
  sessionId?: string;
}

interface SessionEventsResponse {
  sessionId: string;
  events: Array<{ id: string; type: string; payload: string; createdAt: string }>;
}

interface SessionsListResponse {
  sessions: Array<{ id: string; createdAt: string; lastActivityAt: string; messageCount: number }>;
}

interface SessionsSearchResponse {
  sessions: Array<{ sessionId: string; createdAt: string; lastActivityAt: string; messageCount: number; snippet: string }>;
}

interface SessionMessagesResponse {
  sessionId: string;
  messages: Array<{ id: string; role: "user" | "assistant"; content: string; createdAt: string }>;
}

interface TelemetryCountersResponse {
  generatedAt: string;
  counters: {
    newBetaUsers7d: number;
    dailyActiveUsers: number;
    promptsSent: number;
    promptsSent24h: number;
    applyActionsProposed: number;
    applyActionsConfirmed: number;
    blockedCommandAttempts: number;
    streamRequests: number;
    streamCompletions: number;
    streamSuccessRate: number;
    medianChatLatencyMs: number;
    medianStreamFirstTokenLatencyMs: number;
  };
}

interface CliModeArgs {
  runKpi: boolean;
  email?: string;
  jsonMode: boolean;
}

function parsePositiveInt(value: string | undefined, fallback: number, maxValue: number): number {
  if (!value) {
    return fallback;
  }

  const parsed = Number.parseInt(value, 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return fallback;
  }

  return Math.min(parsed, maxValue);
}

function parseArgsWithJsonFlag(parts: string[]): { args: string[]; jsonMode: boolean } {
  const jsonMode = parts.includes("--json");
  const args = parts.filter((item) => item !== "--json");
  return { args, jsonMode };
}

function parseCliModeArgs(parts: string[]): CliModeArgs {
  let runKpi = false;
  let email: string | undefined;
  let jsonMode = false;

  for (let index = 0; index < parts.length; index += 1) {
    const token = parts[index];

    if (token === "--kpi") {
      runKpi = true;
      continue;
    }

    if (token === "--json") {
      jsonMode = true;
      continue;
    }

    if (token === "--email") {
      const value = parts[index + 1]?.trim();
      if (!value || value.startsWith("--")) {
        throw new Error("Usage: npm run cli -- --kpi [--email <email>] [--json]");
      }

      email = value;
      index += 1;
      continue;
    }

    if (token.startsWith("--email=")) {
      const value = token.slice("--email=".length).trim();
      if (!value) {
        throw new Error("Usage: npm run cli -- --kpi [--email <email>] [--json]");
      }

      email = value;
      continue;
    }

    throw new Error(`Unknown option: ${token}`);
  }

  return { runKpi, email, jsonMode };
}

function parseArgsWithExportFlags(parts: string[]): {
  args: string[];
  jsonMode: boolean;
  jsonlMode: boolean;
  noMeta: boolean;
  pretty: boolean;
  stdoutMode: boolean;
  raw: boolean;
  gzip: boolean;
  autoExt: boolean;
  timestampName: boolean;
  mkdir: boolean;
  safeWrite: boolean;
  overwrite: boolean;
  ifExists: "overwrite" | "error" | "skip" | "rename";
  diagnosticsFormat: "text" | "json";
  error?: string;
} {
  let jsonMode = false;
  let jsonlMode = false;
  let noMeta = false;
  let pretty = false;
  let stdoutMode = false;
  let raw = false;
  let gzip = false;
  let autoExt = false;
  let timestampName = false;
  let mkdir = false;
  let safeWrite = false;
  let overwrite = true;
  let ifExists: "overwrite" | "error" | "skip" | "rename" = "overwrite";
  let diagnosticsFormat: "text" | "json" = "text";
  const args: string[] = [];

  for (const item of parts) {
    if (item === "--json") {
      jsonMode = true;
      continue;
    }

    if (item === "--jsonl") {
      jsonlMode = true;
      continue;
    }

    if (item === "--no-meta") {
      noMeta = true;
      continue;
    }

    if (item === "--pretty") {
      pretty = true;
      continue;
    }

    if (item === "--stdout") {
      stdoutMode = true;
      continue;
    }

    if (item === "--raw") {
      raw = true;
      continue;
    }

    if (item === "--gzip") {
      gzip = true;
      continue;
    }

    if (item === "--auto-ext") {
      autoExt = true;
      continue;
    }

    if (item === "--timestamp-name") {
      timestampName = true;
      continue;
    }

    if (item === "--mkdir") {
      mkdir = true;
      continue;
    }

    if (item === "--safe-write") {
      safeWrite = true;
      continue;
    }

    if (item.startsWith("--overwrite=")) {
      const value = item.slice("--overwrite=".length).trim().toLowerCase();
      if (value === "true") {
        overwrite = true;
        ifExists = "overwrite";
        continue;
      }

      if (value === "false") {
        overwrite = false;
        ifExists = "error";
        continue;
      }

      return {
        args,
        jsonMode,
        jsonlMode,
        noMeta,
        pretty,
        stdoutMode,
        raw,
        gzip,
        autoExt,
        timestampName,
        mkdir,
        safeWrite,
        overwrite,
        ifExists,
        diagnosticsFormat,
        error: "Use --overwrite=true or --overwrite=false"
      };
    }

    if (item.startsWith("--if-exists=")) {
      const value = item.slice("--if-exists=".length).trim().toLowerCase();
      if (value === "overwrite" || value === "error" || value === "skip" || value === "rename") {
        ifExists = value;
        overwrite = value === "overwrite";
        continue;
      }

      return {
        args,
        jsonMode,
        jsonlMode,
        noMeta,
        pretty,
        stdoutMode,
        raw,
        gzip,
        autoExt,
        timestampName,
        mkdir,
        safeWrite,
        overwrite,
        ifExists,
        diagnosticsFormat,
        error: "Use --if-exists=overwrite|error|skip|rename"
      };
    }

    if (item.startsWith("--diagnostics=")) {
      const value = item.slice("--diagnostics=".length).trim().toLowerCase();
      if (value === "text" || value === "json") {
        diagnosticsFormat = value;
        continue;
      }

      return {
        args,
        jsonMode,
        jsonlMode,
        noMeta,
        pretty,
        stdoutMode,
        raw,
        gzip,
        autoExt,
        timestampName,
        mkdir,
        safeWrite,
        overwrite,
        ifExists,
        diagnosticsFormat,
        error: "Use --diagnostics=text or --diagnostics=json"
      };
    }

    if (item.startsWith("--")) {
      return {
        args,
        jsonMode,
        jsonlMode,
        noMeta,
        pretty,
        stdoutMode,
        raw,
        gzip,
        autoExt,
        timestampName,
        mkdir,
        safeWrite,
        overwrite,
        ifExists,
        diagnosticsFormat,
        error: `Unknown option: ${item}`
      };
    }

    args.push(item);
  }

  return { args, jsonMode, jsonlMode, noMeta, pretty, stdoutMode, raw, gzip, autoExt, timestampName, mkdir, safeWrite, overwrite, ifExists, diagnosticsFormat };
}

function formatTimestampForFileName(now: Date): string {
  const iso = now.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
  const datePart = iso.slice(0, 8);
  const timePart = iso.slice(9, 15);
  return `${datePart}-${timePart}Z`;
}

function resolveHistoryExportPath(
  outputFile: string,
  options: { jsonMode: boolean; jsonlMode: boolean; gzip: boolean; autoExt: boolean; timestampName: boolean }
): string {
  let resolved = outputFile;
  if (options.timestampName) {
    const parsed = path.parse(resolved);
    const timestamp = formatTimestampForFileName(new Date());
    const baseName = `${parsed.name}-${timestamp}${parsed.ext}`;
    resolved = parsed.dir ? path.join(parsed.dir, baseName) : baseName;
  }

  if (!options.autoExt) {
    return resolved;
  }

  const lower = resolved.toLowerCase();

  if (options.jsonMode && !lower.endsWith(".json") && !lower.endsWith(".json.gz")) {
    resolved += ".json";
  } else if (options.jsonlMode && !lower.endsWith(".jsonl") && !lower.endsWith(".jsonl.gz")) {
    resolved += ".jsonl";
  } else if (!options.jsonMode && !options.jsonlMode && !lower.endsWith(".txt") && !lower.endsWith(".txt.gz")) {
    resolved += ".txt";
  }

  if (options.gzip && !resolved.toLowerCase().endsWith(".gz")) {
    resolved += ".gz";
  }

  return resolved;
}

function resolveNextAvailablePath(filePath: string): string {
  const parsed = path.parse(filePath);
  let sequence = 1;
  while (true) {
    const fileName = `${parsed.name}-${sequence}${parsed.ext}`;
    const candidate = parsed.dir ? path.join(parsed.dir, fileName) : fileName;
    if (!fs.existsSync(candidate)) {
      return candidate;
    }

    sequence += 1;
  }
}

function writeFileSyncSafe(
  filePath: string,
  data: string | Buffer,
  safeWrite: boolean,
  ifExists: "overwrite" | "error" | "skip" | "rename"
): { filePath: string; written: boolean; collisionAction: "none" | "skipped" | "renamed" } {
  let targetPath = filePath;
  if (ifExists === "skip" && fs.existsSync(targetPath)) {
    return { filePath: targetPath, written: false, collisionAction: "skipped" };
  }

  if (ifExists === "error" && fs.existsSync(targetPath)) {
    throw new Error(`Target file already exists: ${targetPath}`);
  }

  let collisionAction: "none" | "renamed" = "none";
  if (ifExists === "rename" && fs.existsSync(targetPath)) {
    targetPath = resolveNextAvailablePath(targetPath);
    collisionAction = "renamed";
  }

  if (!safeWrite) {
    if (typeof data === "string") {
      fs.writeFileSync(targetPath, data, "utf8");
    } else {
      fs.writeFileSync(targetPath, data);
    }

    return { filePath: targetPath, written: true, collisionAction };
  }

  const directory = path.dirname(targetPath);
  const tempPath = path.join(directory, `.${path.basename(targetPath)}.${process.pid}.${Date.now()}.tmp`);

  try {
    if (typeof data === "string") {
      fs.writeFileSync(tempPath, data, "utf8");
    } else {
      fs.writeFileSync(tempPath, data);
    }

    if (fs.existsSync(targetPath)) {
      fs.rmSync(targetPath, { force: true });
    }

    fs.renameSync(tempPath, targetPath);
    return { filePath: targetPath, written: true, collisionAction };
  } catch (error) {
    if (fs.existsSync(tempPath)) {
      fs.rmSync(tempPath, { force: true });
    }

    throw error;
  }
}

function parseQueryAndLimit(parts: string[], fallbackLimit: number, maxLimit: number): { query: string; limit: number; jsonMode: boolean } {
  const { args, jsonMode } = parseArgsWithJsonFlag(parts);
  if (args.length < 2) {
    return { query: "", limit: fallbackLimit, jsonMode };
  }

  let queryTerms = args.slice(1);
  let limit = fallbackLimit;
  const maybeLast = queryTerms[queryTerms.length - 1];

  if (maybeLast && /^\d+$/.test(maybeLast)) {
    limit = parsePositiveInt(maybeLast, fallbackLimit, maxLimit);
    queryTerms = queryTerms.slice(0, -1);
  }

  return {
    query: queryTerms.join(" ").trim(),
    limit,
    jsonMode
  };
}

function buildContextFromMessages(
  messages: Array<{ role: string; content: string; createdAt?: string }>,
  maxChars: number
): { context: string; messageCount: number } {
  const lines = messages.map((item) => {
    const stamp = item.createdAt ? `[${item.createdAt}] ` : "";
    return `${stamp}${item.role}> ${item.content}`;
  });

  return {
    context: lines.join("\n").slice(-maxChars),
    messageCount: messages.length
  };
}

function parseImportedHistoryFile(filePath: string, raw: string, maxChars: number): { context: string; format: "text" | "json" | "jsonl"; messageCount?: number } {
  const extension = path.extname(filePath).toLowerCase();

  if (extension === ".json") {
    const parsed = JSON.parse(raw) as
      | { messages?: Array<{ role?: string; content?: string; createdAt?: string }> }
      | Array<{ role?: string; content?: string; createdAt?: string }>;

    const source = Array.isArray(parsed) ? parsed : parsed.messages ?? [];
    const normalized = source
      .filter((item) => typeof item?.role === "string" && typeof item?.content === "string")
      .map((item) => ({
        role: item.role as string,
        content: item.content as string,
        createdAt: item.createdAt
      }));

    const built = buildContextFromMessages(normalized, maxChars);
    return { context: built.context, format: "json", messageCount: built.messageCount };
  }

  if (extension === ".jsonl") {
    const normalized = raw
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => JSON.parse(line) as { role?: string; content?: string; createdAt?: string })
      .filter((item) => typeof item?.role === "string" && typeof item?.content === "string")
      .map((item) => ({
        role: item.role as string,
        content: item.content as string,
        createdAt: item.createdAt
      }));

    const built = buildContextFromMessages(normalized, maxChars);
    return { context: built.context, format: "jsonl", messageCount: built.messageCount };
  }

  return { context: raw.slice(-maxChars), format: "text" };
}

function parseImportedHistoryAuto(raw: string, maxChars: number): { context: string; format: "text" | "json" | "jsonl"; messageCount?: number } {
  const trimmed = raw.trim();

  if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
    try {
      return parseImportedHistoryFile("stdin.json", raw, maxChars);
    } catch {
      // fall through to jsonl/text checks
    }
  }

  const lines = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length > 0 && lines.every((line) => line.startsWith("{") && line.endsWith("}"))) {
    try {
      return parseImportedHistoryFile("stdin.jsonl", raw, maxChars);
    } catch {
      // fall through to text
    }
  }

  return { context: raw.slice(-maxChars), format: "text" };
}

function parseImportedHistoryFromFile(filePath: string, maxChars: number): {
  parsed: { context: string; format: "text" | "json" | "jsonl"; messageCount?: number };
  inputChars: number;
  compressed: boolean;
} {
  const fileBuffer = fs.readFileSync(filePath);
  const lowerFilePath = filePath.toLowerCase();
  const isGzip = lowerFilePath.endsWith(".gz");
  const decoded = isGzip ? gunzipSync(fileBuffer).toString("utf8") : fileBuffer.toString("utf8");

  const parserFilePath = isGzip ? filePath.slice(0, -3) : filePath;
  const parserExt = path.extname(parserFilePath).toLowerCase();
  const parsed = parserExt === ".json" || parserExt === ".jsonl"
    ? parseImportedHistoryFile(parserFilePath, decoded, maxChars)
    : parseImportedHistoryAuto(decoded, maxChars);

  return {
    parsed,
    inputChars: decoded.length,
    compressed: isGzip
  };
}

function parseImportHistoryArgs(parts: string[]): { inputSource?: string; maxContext: number; pretty: boolean; error?: string } {
  const args = parts.slice(1);
  if (args.length === 0) {
    return { maxContext: 1200, pretty: false, error: "Usage: /import-history <filePath|-> [--max-context <chars>] [--pretty]" };
  }

  let inputSource: string | undefined;
  let maxContext = 1200;
  let pretty = false;

  for (let index = 0; index < args.length; index += 1) {
    const token = args[index];

    if (token === "--max-context") {
      const next = args[index + 1];
      const parsed = Number.parseInt(next ?? "", 10);
      if (!Number.isFinite(parsed) || parsed <= 0) {
        return { maxContext, pretty, error: "Usage: /import-history <filePath|-> [--max-context <chars>] [--pretty]" };
      }

      maxContext = Math.min(parsed, 20000);
      index += 1;
      continue;
    }

    if (token === "--pretty") {
      pretty = true;
      continue;
    }

    if (token.startsWith("--")) {
      return { maxContext, pretty, error: `Unknown option: ${token}` };
    }

    if (inputSource) {
      return { maxContext, pretty, error: "Provide only one input source: <filePath> or -" };
    }

    inputSource = token;
  }

  if (!inputSource) {
    return { maxContext, pretty, error: "Usage: /import-history <filePath|-> [--max-context <chars>] [--pretty]" };
  }

  return { inputSource, maxContext, pretty };
}

async function streamChat(
  baseUrl: string,
  token: string,
  sessionId: string | undefined,
  message: string,
  cwd: string,
  lastOutput: string
) {
  const params = new URLSearchParams({ message, cwd, lastOutput });
  if (sessionId) {
    params.set("sessionId", sessionId);
  }

  const response = await fetch(`${baseUrl}/chat/stream?${params.toString()}`, {
    headers: {
      authorization: `Bearer ${token}`
    }
  });
  if (!response.ok || !response.body) {
    throw new Error(`stream failed: ${response.status}`);
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();

  let buffer = "";
  let currentSessionId = sessionId;
  let proposal: unknown = null;

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }

    buffer += decoder.decode(value, { stream: true });
    const events = buffer.split("\n\n");
    buffer = events.pop() ?? "";

    for (const event of events) {
      if (!event.startsWith("data: ")) {
        continue;
      }

      const parsed = JSON.parse(event.slice(6)) as ChatChunkEvent;

      if (parsed.sessionId) {
        currentSessionId = parsed.sessionId;
      }

      if (parsed.type === "chunk") {
        process.stdout.write(String(parsed.data ?? ""));
      }

      if (parsed.type === "proposal") {
        proposal = parsed.data;
      }

      if (parsed.type === "done") {
        process.stdout.write("\n");
      }
    }
  }

  return { sessionId: currentSessionId, proposal };
}

async function devLogin(baseUrl: string, email: string): Promise<DevLoginResponse> {
  const loginResponse = await fetch(`${baseUrl}/auth/dev-login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ email })
  });

  if (!loginResponse.ok) {
    throw new Error(`dev-login failed: ${loginResponse.status}`);
  }

  return (await loginResponse.json()) as DevLoginResponse;
}

async function fetchTelemetryCounters(baseUrl: string, token: string): Promise<TelemetryCountersResponse> {
  const telemetryResponse = await fetch(`${baseUrl}/telemetry/counters`, {
    headers: {
      authorization: `Bearer ${token}`
    }
  });

  if (!telemetryResponse.ok) {
    throw new Error(`Failed to fetch KPI counters: ${telemetryResponse.status}`);
  }

  return (await telemetryResponse.json()) as TelemetryCountersResponse;
}

function formatKpi(telemetryData: TelemetryCountersResponse, jsonMode: boolean): string {
  if (jsonMode) {
    return `${JSON.stringify(telemetryData, null, 2)}\n`;
  }

  const counters = telemetryData.counters;
  const streamSuccessPct = `${(counters.streamSuccessRate * 100).toFixed(1)}%`;

  return [
    `KPI counters @ ${telemetryData.generatedAt}`,
    `- New beta users (7d): ${counters.newBetaUsers7d}`,
    `- Daily active users: ${counters.dailyActiveUsers}`,
    `- Prompts sent (all-time): ${counters.promptsSent}`,
    `- Prompts sent (24h): ${counters.promptsSent24h}`,
    `- Apply actions proposed: ${counters.applyActionsProposed}`,
    `- Apply actions confirmed: ${counters.applyActionsConfirmed}`,
    `- Blocked command attempts: ${counters.blockedCommandAttempts}`,
    `- Stream requests: ${counters.streamRequests}`,
    `- Stream completions: ${counters.streamCompletions}`,
    `- Stream success rate: ${streamSuccessPct}`,
    `- Median /chat latency: ${counters.medianChatLatencyMs.toFixed(0)} ms`,
    `- Median /chat/stream first-token latency: ${counters.medianStreamFirstTokenLatencyMs.toFixed(0)} ms`,
    ""
  ].join("\n");
}

async function main() {
  const mode = parseCliModeArgs(process.argv.slice(2));
  const baseUrl = process.env.TSIFULATOR_API_URL ?? "http://localhost:4000";

  if (mode.runKpi) {
    const email = mode.email?.trim() || "founder@tsifulator.ai";
    const login = await devLogin(baseUrl, email);
    const telemetryData = await fetchTelemetryCounters(baseUrl, login.token);
    output.write(formatKpi(telemetryData, mode.jsonMode));
    return;
  }

  const rl = readline.createInterface({ input, output });
  const email = (await rl.question("Dev email: ")).trim() || "founder@tsifulator.ai";

  const login = await devLogin(baseUrl, email);

  output.write(`Logged in as ${login.user.email}\n`);
  output.write("Type your prompt. Use 'cmd: <command>' to request a command proposal.\n");
  output.write("Commands: /help, /session, /session-new, /session-clone <id> [limit], /context, /context-clear, /events, /kpi [--json], /sessions [limit] [--json], /sessions-find <text> [limit] [--json], /history [sessionId] [limit] [--json], /history-export <file> [sessionId] [limit] [--json|--jsonl] [--no-meta] [--pretty] [--stdout] [--raw] [--gzip] [--auto-ext] [--timestamp-name] [--mkdir] [--safe-write] [--overwrite=true|false] [--if-exists=overwrite|error|skip|rename] [--diagnostics=text|json], /import-history <file|-> [--max-context <chars>] [--pretty], /feedback <text>, /use <id>, /clear, exit\n\n");

  let sessionId: string | undefined;
  let lastOutput = "";

  while (true) {
    const prompt = (await rl.question("tsifulator> ")).trim();
    if (!prompt || prompt.toLowerCase() === "exit") {
      break;
    }

    if (prompt === "/help") {
      output.write("/help    Show commands\n");
      output.write("/session Show current session id\n");
      output.write("/session-new Start a fresh session\n");
      output.write("/session-clone <id> [limit] Clone history into new session context\n");
      output.write("/context Show active context summary\n");
      output.write("/context-clear Clear loaded context only\n");
      output.write("/events  Show latest event count for current session\n\n");
      output.write("/kpi [--json] Show telemetry counters dashboard\n\n");
      output.write("/sessions [limit] [--json] List your recent sessions\n");
      output.write("/sessions-find <text> [limit] [--json] Search sessions by message content\n");
      output.write("/history [sessionId] [limit] [--json] Show session messages\n\n");
      output.write("/history-export <file> [sessionId] [limit] [--json|--jsonl] [--no-meta] [--pretty] [--stdout] [--raw] [--gzip] [--auto-ext] [--timestamp-name] [--mkdir] [--safe-write] [--overwrite=true|false] [--if-exists=overwrite|error|skip|rename] [--diagnostics=text|json] Save history to file or emit to stdout\n\n");
      output.write("/import-history <file|-> [--max-context <chars>] [--pretty] Load file or pasted stdin content into context for next prompts\n\n");
      output.write("/feedback <text> Capture product feedback for this session\n\n");
      output.write("/use <id> Switch to an existing session\n");
      output.write("/clear    Clear active session selection\n\n");
      continue;
    }

    if (prompt === "/session") {
      output.write(`Current session: ${sessionId ?? "(none yet)"}\n\n`);
      continue;
    }

    if (prompt === "/session-new") {
      sessionId = undefined;
      output.write("Fresh session armed. Your next prompt will create a new session id.\n\n");
      continue;
    }

    if (prompt.startsWith("/session-clone ")) {
      const parts = prompt.split(/\s+/).filter(Boolean);
      const sourceSessionId = parts[1];
      const limit = parsePositiveInt(parts[2], 40, 500);

      if (!sourceSessionId) {
        output.write("Usage: /session-clone <sessionId> [limit]\n\n");
        continue;
      }

      const historyResponse = await fetch(`${baseUrl}/sessions/${sourceSessionId}/messages?limit=${limit}`, {
        headers: {
          authorization: `Bearer ${login.token}`
        }
      });

      if (!historyResponse.ok) {
        output.write(`Cannot clone session ${sourceSessionId}: ${historyResponse.status}\n\n`);
        continue;
      }

      const historyData = (await historyResponse.json()) as SessionMessagesResponse;
      const snapshot = historyData.messages.map((item) => `${item.role}> ${item.content}`).join("\n");

      lastOutput = snapshot.slice(-1200);
      sessionId = undefined;

      output.write(`Cloned context from ${sourceSessionId} (${historyData.messages.length} messages).\n`);
      output.write("Fresh session armed. Next prompt will start a new session using this context.\n\n");
      continue;
    }

    if (prompt === "/context") {
      const preview = lastOutput ? lastOutput.slice(0, 220).replace(/\s+/g, " ") : "(none)";
      output.write(`Current session: ${sessionId ?? "(none yet)"}\n`);
      output.write(`Context chars: ${lastOutput.length}\n`);
      output.write(`Context preview: ${preview}\n\n`);
      continue;
    }

    if (prompt === "/context-clear") {
      lastOutput = "";
      output.write("Context cleared. Active session remains unchanged.\n\n");
      continue;
    }

    if (prompt === "/events") {
      if (!sessionId) {
        output.write("No active session yet. Send a prompt first.\n\n");
        continue;
      }

      const eventsResponse = await fetch(`${baseUrl}/sessions/${sessionId}/events`, {
        headers: {
          authorization: `Bearer ${login.token}`
        }
      });

      if (!eventsResponse.ok) {
        output.write(`Failed to fetch events: ${eventsResponse.status}\n\n`);
        continue;
      }

      const eventsData = (await eventsResponse.json()) as SessionEventsResponse;
      output.write(`Events in session ${eventsData.sessionId}: ${eventsData.events.length}\n\n`);
      continue;
    }

    if (prompt === "/kpi" || prompt.startsWith("/kpi ")) {
      const parts = prompt.split(/\s+/).filter(Boolean);
      const { jsonMode } = parseArgsWithJsonFlag(parts);

      try {
        const telemetryData = await fetchTelemetryCounters(baseUrl, login.token);
        output.write(formatKpi(telemetryData, jsonMode));
      } catch (error) {
        output.write(`${error instanceof Error ? error.message : String(error)}\n\n`);
      }

      continue;
    }

    if (prompt === "/sessions" || prompt.startsWith("/sessions ")) {
      const parts = prompt.split(/\s+/).filter(Boolean);
      const { args, jsonMode } = parseArgsWithJsonFlag(parts);
      const limit = parsePositiveInt(args[1], 10, 100);

      const sessionsResponse = await fetch(`${baseUrl}/sessions?limit=${limit}`, {
        headers: {
          authorization: `Bearer ${login.token}`
        }
      });

      if (!sessionsResponse.ok) {
        output.write(`Failed to fetch sessions: ${sessionsResponse.status}\n\n`);
        continue;
      }

      const sessionsData = (await sessionsResponse.json()) as SessionsListResponse;
      if (sessionsData.sessions.length === 0) {
        output.write("No sessions found yet.\n\n");
        continue;
      }

      if (jsonMode) {
        output.write(`${JSON.stringify(sessionsData.sessions, null, 2)}\n\n`);
        continue;
      }

      output.write(`Recent sessions (limit=${limit}):\n`);
      for (const item of sessionsData.sessions) {
        output.write(`- ${item.id} | messages=${item.messageCount} | last=${item.lastActivityAt}\n`);
      }
      output.write("\n");
      continue;
    }

    if (prompt.startsWith("/sessions-find")) {
      const parts = prompt.split(/\s+/).filter(Boolean);
      const parsed = parseQueryAndLimit(parts, 10, 100);
      if (!parsed.query) {
        output.write("Usage: /sessions-find <text> [limit]\n\n");
        continue;
      }

      const q = parsed.query;
      const limit = parsed.limit;
      const jsonMode = parsed.jsonMode;

      const params = new URLSearchParams({ q, limit: String(limit) });
      const searchResponse = await fetch(`${baseUrl}/sessions/search?${params.toString()}`, {
        headers: {
          authorization: `Bearer ${login.token}`
        }
      });

      if (!searchResponse.ok) {
        output.write(`Failed to search sessions: ${searchResponse.status}\n\n`);
        continue;
      }

      const searchData = (await searchResponse.json()) as SessionsSearchResponse;
      if (searchData.sessions.length === 0) {
        output.write(`No sessions matched '${q}'.\n\n`);
        continue;
      }

      if (jsonMode) {
        output.write(`${JSON.stringify(searchData.sessions, null, 2)}\n\n`);
        continue;
      }

      output.write(`Session matches for '${q}' (limit=${limit}):\n`);
      for (const item of searchData.sessions) {
        output.write(`- ${item.sessionId} | messages=${item.messageCount} | last=${item.lastActivityAt}\n`);
        output.write(`  snippet: ${item.snippet}\n`);
      }
      output.write("\n");
      continue;
    }

    if (prompt.startsWith("/use ")) {
      const targetSessionId = prompt.slice(5).trim();
      if (!targetSessionId) {
        output.write("Usage: /use <sessionId>\n\n");
        continue;
      }

      const validateResponse = await fetch(`${baseUrl}/sessions/${targetSessionId}/messages?limit=1`, {
        headers: {
          authorization: `Bearer ${login.token}`
        }
      });

      if (!validateResponse.ok) {
        output.write(`Cannot use session ${targetSessionId}: ${validateResponse.status}\n\n`);
        continue;
      }

      sessionId = targetSessionId;
      output.write(`Switched to session: ${sessionId}\n\n`);
      continue;
    }

    if (prompt === "/clear") {
      sessionId = undefined;
      output.write("Active session cleared. Next prompt will start a new session.\n\n");
      continue;
    }

    if (prompt === "/history" || prompt.startsWith("/history ")) {
      const parts = prompt.split(/\s+/).filter(Boolean);
      let targetSessionId = sessionId;
      let limit = 12;
      const { args, jsonMode } = parseArgsWithJsonFlag(parts);

      if (args.length >= 2) {
        if (/^\d+$/.test(args[1])) {
          limit = parsePositiveInt(args[1], 12, 500);
        } else {
          targetSessionId = args[1];
        }
      }

      if (args.length >= 3) {
        limit = parsePositiveInt(args[2], 12, 500);
      }

      if (!targetSessionId) {
        output.write("No active session yet. Send a prompt first.\n\n");
        continue;
      }

      const historyResponse = await fetch(`${baseUrl}/sessions/${targetSessionId}/messages?limit=${limit}`, {
        headers: {
          authorization: `Bearer ${login.token}`
        }
      });

      if (!historyResponse.ok) {
        output.write(`Failed to fetch history: ${historyResponse.status}\n\n`);
        continue;
      }

      const historyData = (await historyResponse.json()) as SessionMessagesResponse;

      if (jsonMode) {
        output.write(`${JSON.stringify(historyData.messages, null, 2)}\n\n`);
        continue;
      }

      output.write(`History for ${historyData.sessionId} (limit=${limit}):\n`);
      for (const item of historyData.messages) {
        output.write(`${item.role}> ${item.content}\n`);
      }
      output.write("\n");
      continue;
    }

    if (prompt.startsWith("/history-export")) {
      const parts = prompt.split(/\s+/).filter(Boolean);
      const { args, jsonMode, jsonlMode, noMeta, pretty, stdoutMode, raw, gzip, autoExt, timestampName, mkdir, safeWrite, overwrite, ifExists, diagnosticsFormat, error } = parseArgsWithExportFlags(parts);

      if (error) {
        output.write(`${error}\n\n`);
        continue;
      }

      const positional = args.slice(1);

      if (!stdoutMode && positional.length < 1) {
        output.write("Usage: /history-export <filePath> [sessionId] [limit] [--json|--jsonl] [--no-meta] [--pretty] [--stdout] [--raw] [--gzip] [--auto-ext] [--timestamp-name] [--mkdir] [--safe-write] [--overwrite=true|false] [--if-exists=overwrite|error|skip|rename] [--diagnostics=text|json]\n\n");
        continue;
      }

      if (raw && !stdoutMode) {
        output.write("--raw is only supported with --stdout\n\n");
        continue;
      }

      if (gzip && stdoutMode) {
        output.write("--gzip is only supported for file output (without --stdout)\n\n");
        continue;
      }

      if (timestampName && stdoutMode) {
        output.write("--timestamp-name is only supported for file output (without --stdout)\n\n");
        continue;
      }

      if (mkdir && stdoutMode) {
        output.write("--mkdir is only supported for file output (without --stdout)\n\n");
        continue;
      }

      if (safeWrite && stdoutMode) {
        output.write("--safe-write is only supported for file output (without --stdout)\n\n");
        continue;
      }

      if (!overwrite && stdoutMode) {
        output.write("--overwrite=false is only supported for file output (without --stdout)\n\n");
        continue;
      }

      if (ifExists !== "overwrite" && stdoutMode) {
        output.write("--if-exists is only supported for file output (without --stdout)\n\n");
        continue;
      }

      if (jsonMode && jsonlMode) {
        output.write("Use only one export format flag: --json or --jsonl\n\n");
        continue;
      }

      let targetSessionId = sessionId;
      let limit = 100;
      let outputFile: string | undefined;
      let requestedPath = "stdout";
      let index = 0;

      if (!stdoutMode) {
        requestedPath = positional[0];
        outputFile = resolveHistoryExportPath(positional[0], { jsonMode, jsonlMode, gzip, autoExt, timestampName });
        index = 1;
      }

      const resolvedOutputPath = !stdoutMode ? path.resolve(outputFile as string) : undefined;
      let createdDirs = false;
      if (resolvedOutputPath && mkdir) {
        const targetDirectory = path.dirname(resolvedOutputPath);
        const directoryExisted = fs.existsSync(targetDirectory);
        fs.mkdirSync(targetDirectory, { recursive: true });
        createdDirs = !directoryExisted && fs.existsSync(targetDirectory);
      }

      if (positional.length > index) {
        if (/^\d+$/.test(positional[index])) {
          limit = parsePositiveInt(positional[index], 100, 500);
        } else {
          targetSessionId = positional[index];
          index += 1;
        }
      }

      if (positional.length > index) {
        limit = parsePositiveInt(positional[index], 100, 500);
        index += 1;
      }

      if (positional.length > index) {
        output.write("Too many positional arguments for /history-export\n\n");
        continue;
      }

      if (!targetSessionId) {
        output.write("No active session yet. Provide a sessionId or send a prompt first.\n\n");
        continue;
      }

      const historyResponse = await fetch(`${baseUrl}/sessions/${targetSessionId}/messages?limit=${limit}`, {
        headers: {
          authorization: `Bearer ${login.token}`
        }
      });

      if (!historyResponse.ok) {
        output.write(`Failed to fetch history for export: ${historyResponse.status}\n\n`);
        continue;
      }

      const historyData = (await historyResponse.json()) as SessionMessagesResponse;
      const exportedAt = new Date().toISOString();

      if (jsonMode) {
        const exportData = noMeta
          ? historyData.messages
          : {
              sessionId: historyData.sessionId,
              exportedAt,
              messageCount: historyData.messages.length,
              messages: historyData.messages
            };

        const payload = JSON.stringify(exportData, null, 2);
        const payloadBytes = Buffer.byteLength(payload, "utf8");
        let bytesOut = payloadBytes;
        let effectivePath = stdoutMode ? "stdout" : (resolvedOutputPath as string);
        let collisionAction: "none" | "skipped" | "renamed" = "none";
        if (stdoutMode) {
          output.write(`${payload}\n`);
        } else {
          const filePolicy = ifExists ?? (overwrite ? "overwrite" : "error");
          let writeResult: { filePath: string; written: boolean; collisionAction: "none" | "skipped" | "renamed" };
          if (gzip) {
            const compressed = gzipSync(payload);
            writeResult = writeFileSyncSafe(resolvedOutputPath as string, compressed, safeWrite, filePolicy);
            bytesOut = compressed.length;
          } else {
            writeResult = writeFileSyncSafe(resolvedOutputPath as string, payload, safeWrite, filePolicy);
          }

          if (!writeResult.written) {
            if (!raw) {
              output.write(`History export skipped (exists): ${writeResult.filePath}\n\n`);
            }
            continue;
          }

          effectivePath = writeResult.filePath;
          collisionAction = writeResult.collisionAction;

          if (!raw) {
            output.write(`History exported (json${gzip ? "+gzip" : ""}) to: ${writeResult.filePath}\n`);
          }
        }
        if (pretty && !raw) {
          const pathChanged = requestedPath !== effectivePath;
          const diagnostics = {
            destination: stdoutMode ? "stdout" : "file",
            writeMode: safeWrite ? "safe-write" : "direct",
            format: "json",
            ifExists,
            collisionAction,
            requestedPath,
            effectivePath,
            pathChanged,
            createdDirs,
            noMeta,
            gzip,
            messages: historyData.messages.length,
            bytesIn: payloadBytes,
            bytesOut
          };
          if (diagnosticsFormat === "json") {
            output.write(`${JSON.stringify(diagnostics)}\n`);
          } else {
            output.write(`Export diagnostics: destination=${diagnostics.destination}, writeMode=${diagnostics.writeMode}, format=${diagnostics.format}, ifExists=${diagnostics.ifExists}, collisionAction=${diagnostics.collisionAction}, requestedPath=${diagnostics.requestedPath}, effectivePath=${diagnostics.effectivePath}, pathChanged=${diagnostics.pathChanged}, createdDirs=${diagnostics.createdDirs}, noMeta=${diagnostics.noMeta}, gzip=${diagnostics.gzip}, messages=${diagnostics.messages}, bytesIn=${diagnostics.bytesIn}, bytesOut=${diagnostics.bytesOut}\n`);
          }
        }
        if (!raw) {
          output.write("\n");
        }
        continue;
      }

      if (jsonlMode) {
        const lines = historyData.messages.map((item) =>
          noMeta
            ? JSON.stringify(item)
            : JSON.stringify({
                sessionId: historyData.sessionId,
                exportedAt,
                messageCount: historyData.messages.length,
                ...item
              })
        );
        const payload = lines.join("\n");
        const payloadBytes = Buffer.byteLength(payload, "utf8");
        let bytesOut = payloadBytes;
        let effectivePath = stdoutMode ? "stdout" : (resolvedOutputPath as string);
        let collisionAction: "none" | "skipped" | "renamed" = "none";
        if (stdoutMode) {
          output.write(`${payload}\n`);
        } else {
          const filePolicy = ifExists ?? (overwrite ? "overwrite" : "error");
          let writeResult: { filePath: string; written: boolean; collisionAction: "none" | "skipped" | "renamed" };
          if (gzip) {
            const compressed = gzipSync(payload);
            writeResult = writeFileSyncSafe(resolvedOutputPath as string, compressed, safeWrite, filePolicy);
            bytesOut = compressed.length;
          } else {
            writeResult = writeFileSyncSafe(resolvedOutputPath as string, payload, safeWrite, filePolicy);
          }

          if (!writeResult.written) {
            if (!raw) {
              output.write(`History export skipped (exists): ${writeResult.filePath}\n\n`);
            }
            continue;
          }

          effectivePath = writeResult.filePath;
          collisionAction = writeResult.collisionAction;

          if (!raw) {
            output.write(`History exported (jsonl${gzip ? "+gzip" : ""}) to: ${writeResult.filePath}\n`);
          }
        }
        if (pretty && !raw) {
          const pathChanged = requestedPath !== effectivePath;
          const diagnostics = {
            destination: stdoutMode ? "stdout" : "file",
            writeMode: safeWrite ? "safe-write" : "direct",
            format: "jsonl",
            ifExists,
            collisionAction,
            requestedPath,
            effectivePath,
            pathChanged,
            createdDirs,
            noMeta,
            gzip,
            messages: historyData.messages.length,
            bytesIn: payloadBytes,
            bytesOut
          };
          if (diagnosticsFormat === "json") {
            output.write(`${JSON.stringify(diagnostics)}\n`);
          } else {
            output.write(`Export diagnostics: destination=${diagnostics.destination}, writeMode=${diagnostics.writeMode}, format=${diagnostics.format}, ifExists=${diagnostics.ifExists}, collisionAction=${diagnostics.collisionAction}, requestedPath=${diagnostics.requestedPath}, effectivePath=${diagnostics.effectivePath}, pathChanged=${diagnostics.pathChanged}, createdDirs=${diagnostics.createdDirs}, noMeta=${diagnostics.noMeta}, gzip=${diagnostics.gzip}, messages=${diagnostics.messages}, bytesIn=${diagnostics.bytesIn}, bytesOut=${diagnostics.bytesOut}\n`);
          }
        }
        if (!raw) {
          output.write("\n");
        }
        continue;
      }

      const exportLines = noMeta
        ? []
        : [
            `sessionId=${historyData.sessionId}`,
            `exportedAt=${exportedAt}`,
            `messageCount=${historyData.messages.length}`,
            ""
          ];

      for (const item of historyData.messages) {
        exportLines.push(`[${item.createdAt}] ${item.role}> ${item.content}`);
      }

      const payload = exportLines.join("\n");
      const payloadBytes = Buffer.byteLength(payload, "utf8");
      let bytesOut = payloadBytes;
      let effectivePath = stdoutMode ? "stdout" : (resolvedOutputPath as string);
      let collisionAction: "none" | "skipped" | "renamed" = "none";
      if (stdoutMode) {
        output.write(`${payload}\n`);
      } else {
        const filePolicy = ifExists ?? (overwrite ? "overwrite" : "error");
        let writeResult: { filePath: string; written: boolean; collisionAction: "none" | "skipped" | "renamed" };
        if (gzip) {
          const compressed = gzipSync(payload);
          writeResult = writeFileSyncSafe(resolvedOutputPath as string, compressed, safeWrite, filePolicy);
          bytesOut = compressed.length;
        } else {
          writeResult = writeFileSyncSafe(resolvedOutputPath as string, payload, safeWrite, filePolicy);
        }

        if (!writeResult.written) {
          if (!raw) {
            output.write(`History export skipped (exists): ${writeResult.filePath}\n\n`);
          }
          continue;
        }

        effectivePath = writeResult.filePath;
        collisionAction = writeResult.collisionAction;

        if (!raw) {
          output.write(`History exported (${gzip ? "gzip" : "text"}) to: ${writeResult.filePath}\n`);
        }
      }
      if (pretty && !raw) {
        const pathChanged = requestedPath !== effectivePath;
        const diagnostics = {
          destination: stdoutMode ? "stdout" : "file",
          writeMode: safeWrite ? "safe-write" : "direct",
          format: "text",
          ifExists,
          collisionAction,
          requestedPath,
          effectivePath,
          pathChanged,
          createdDirs,
          noMeta,
          gzip,
          messages: historyData.messages.length,
          bytesIn: payloadBytes,
          bytesOut
        };
        if (diagnosticsFormat === "json") {
          output.write(`${JSON.stringify(diagnostics)}\n`);
        } else {
          output.write(`Export diagnostics: destination=${diagnostics.destination}, writeMode=${diagnostics.writeMode}, format=${diagnostics.format}, ifExists=${diagnostics.ifExists}, collisionAction=${diagnostics.collisionAction}, requestedPath=${diagnostics.requestedPath}, effectivePath=${diagnostics.effectivePath}, pathChanged=${diagnostics.pathChanged}, createdDirs=${diagnostics.createdDirs}, noMeta=${diagnostics.noMeta}, gzip=${diagnostics.gzip}, messages=${diagnostics.messages}, bytesIn=${diagnostics.bytesIn}, bytesOut=${diagnostics.bytesOut}\n`);
        }
      }
      if (!raw) {
        output.write("\n");
      }
      continue;
    }

    if (prompt.startsWith("/import-history")) {
      const parts = prompt.split(/\s+/).filter(Boolean);
      const parsedArgs = parseImportHistoryArgs(parts);
      if (parsedArgs.error) {
        output.write(`${parsedArgs.error}\n\n`);
        continue;
      }

      const inputSource = parsedArgs.inputSource as string;
      const maxContext = parsedArgs.maxContext;
      const pretty = parsedArgs.pretty;

      if (inputSource === "-") {
        output.write("Paste history content below. End input with a single line: /end\n");

        const pastedLines: string[] = [];
        while (true) {
          const line = await rl.question("");
          if (line.trim() === "/end") {
            break;
          }

          pastedLines.push(line);
        }

        const imported = pastedLines.join("\n");
        if (!imported.trim()) {
          output.write("No input received from stdin paste mode.\n\n");
          continue;
        }

        try {
          const parsedImport = parseImportedHistoryAuto(imported, maxContext);
          lastOutput = parsedImport.context;
          const countSuffix = parsedImport.messageCount !== undefined ? ` (${parsedImport.messageCount} messages)` : "";
          output.write(`Imported ${parsedImport.format} history context from stdin${countSuffix}\n`);
          if (pretty) {
            output.write(`Import diagnostics: source=stdin, format=${parsedImport.format}, maxContext=${maxContext}, inputChars=${imported.length}, loadedChars=${lastOutput.length}\n`);
          }
          output.write(`Context size loaded: ${lastOutput.length} chars\n\n`);
        } catch (error) {
          output.write(`Failed to parse pasted history: ${error instanceof Error ? error.message : String(error)}\n\n`);
        }

        continue;
      }

      const inputFile = path.resolve(inputSource);
      if (!fs.existsSync(inputFile)) {
        output.write(`File not found: ${inputFile}\n\n`);
        continue;
      }

      try {
        const fileImport = parseImportedHistoryFromFile(inputFile, maxContext);
        const parsedImport = fileImport.parsed;
        lastOutput = parsedImport.context;

        const countSuffix = parsedImport.messageCount !== undefined ? ` (${parsedImport.messageCount} messages)` : "";
        output.write(`Imported ${parsedImport.format}${fileImport.compressed ? "+gzip" : ""} history context from: ${inputFile}${countSuffix}\n`);
        if (pretty) {
          output.write(`Import diagnostics: source=file, format=${parsedImport.format}, compressed=${fileImport.compressed}, maxContext=${maxContext}, inputChars=${fileImport.inputChars}, loadedChars=${lastOutput.length}\n`);
        }
        output.write(`Context size loaded: ${lastOutput.length} chars\n\n`);
      } catch (error) {
        output.write(`Failed to parse history file: ${error instanceof Error ? error.message : String(error)}\n\n`);
      }
      continue;
    }

    if (prompt.startsWith("/feedback")) {
      const text = prompt.slice("/feedback".length).trim();
      if (!text) {
        output.write("Usage: /feedback <text>\n\n");
        continue;
      }

      const feedbackResponse = await fetch(`${baseUrl}/feedback`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          authorization: `Bearer ${login.token}`
        },
        body: JSON.stringify({
          sessionId,
          text
        })
      });

      if (!feedbackResponse.ok) {
        output.write(`Feedback failed: ${feedbackResponse.status}\n\n`);
        continue;
      }

      const feedbackData = (await feedbackResponse.json()) as { status: string; sessionId: string };
      sessionId = feedbackData.sessionId;
      output.write(`Feedback saved (session: ${feedbackData.sessionId})\n\n`);
      continue;
    }

    const result = await streamChat(baseUrl, login.token, sessionId, prompt, process.cwd(), lastOutput);
    sessionId = result.sessionId;

    if (result.proposal) {
      const proposal = result.proposal as { id: string; command: string; risk: string; description?: string };
      const riskLabel = proposal.risk === "blocked" ? "🚫 BLOCKED" : proposal.risk === "confirm" ? "⚠️  CONFIRM" : "✅ SAFE";

      output.write("\n┌─── Action Proposal ───────────────────────\n");
      output.write(`│ Command:  ${proposal.command}\n`);
      output.write(`│ Risk:     ${riskLabel}\n`);
      if (proposal.description) {
        output.write(`│ Info:     ${proposal.description}\n`);
      }
      output.write("└────────────────────────────────────────────\n");

      if (proposal.risk === "blocked") {
        output.write("This command is blocked by policy and cannot be executed.\n");
      } else {
        const approve = (await rl.question("Approve action? (y/N): ")).trim().toLowerCase();

        if (approve === "y") {
          const execResponse = await fetch(`${baseUrl}/actions/approve`, {
            method: "POST",
            headers: {
              "content-type": "application/json",
              authorization: `Bearer ${login.token}`
            },
            body: JSON.stringify({
              proposalId: proposal.id,
              approved: true,
              cwd: process.cwd()
            })
          });

          const execution = await execResponse.json() as { status: string; output?: string; hint?: string };
          if (execution.status === "ok") {
            output.write(`\n✅ Command executed successfully\n`);
            if (execution.output) {
              output.write(`\n${execution.output}\n`);
            }
          } else if (execution.status === "blocked") {
            output.write(`\n🚫 Command blocked by policy\n`);
          } else {
            output.write(`\n❌ Command failed (${execution.status})\n`);
            if (execution.output) {
              output.write(`${execution.output}\n`);
            }
            if (execution.hint) {
              output.write(`💡 ${execution.hint}\n`);
            }
          }
          lastOutput = JSON.stringify(execution).slice(0, 1200);
        } else {
          await fetch(`${baseUrl}/actions/approve`, {
            method: "POST",
            headers: {
              "content-type": "application/json",
              authorization: `Bearer ${login.token}`
            },
            body: JSON.stringify({
              proposalId: proposal.id,
              approved: false,
              cwd: process.cwd()
            })
          });
          output.write("↩ Action rejected.\n");
          lastOutput = "action rejected by user";
        }
      }
    }

    output.write("\n");
  }

  rl.close();
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
