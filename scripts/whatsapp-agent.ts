import "dotenv/config";
import Fastify from "fastify";
import formbody from "@fastify/formbody";
import twilio from "twilio";
import { spawn } from "node:child_process";
import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

async function main() {
  const accountSid = process.env.TWILIO_ACCOUNT_SID;
  const authToken = process.env.TWILIO_AUTH_TOKEN;
  const whatsappFrom = process.env.TWILIO_WHATSAPP_FROM;
  const whatsappTo = process.env.YOUR_WHATSAPP_TO;
  const port = Number(process.env.WHATSAPP_AGENT_PORT ?? "8787");

  if (!accountSid || !authToken || !whatsappFrom || !whatsappTo) {
    throw new Error("Missing required env vars: TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_WHATSAPP_FROM, YOUR_WHATSAPP_TO");
  }

  const client = twilio(accountSid, authToken);
  const app = Fastify({ logger: true });

  await app.register(formbody);

  const repoRoot = process.cwd();
  const taskFile = path.join(repoRoot, "current-task.txt");
  const runnerFile = path.join(repoRoot, "run-tsif-agent-once.ps1");
  const ollamaAnswerFile = path.join(repoRoot, "ollama-last-answer.txt");
  const logDir = path.join(repoRoot, "logs");
  const runLog = path.join(logDir, "whatsapp-agent.log");

  if (!existsSync(logDir)) {
    mkdirSync(logDir, { recursive: true });
  }

  let isRunning = false;
  let lastTask = "";

  function logLine(line: string) {
    const stamped = `[${new Date().toISOString()}] ${line}\n`;
    appendFileSync(runLog, stamped, "utf8");
  }

  async function sendWhatsapp(body: string) {
    await client.messages.create({
      from: whatsappFrom,
      to: whatsappTo,
      body
    });
  }

  function normalizePhone(value: string): string {
    return value.replace(/\s+/g, "").trim().toLowerCase();
  }

  function getResultText(): string {
    if (!existsSync(ollamaAnswerFile)) {
      return "No ollama-last-answer.txt file was found.";
    }

    let text = readFileSync(ollamaAnswerFile, "utf8");

    text = text
      .replace(/[^\x09\x0A\x0D\x20-\x7E]/g, "")
      .replace(/\s+/g, " ")
      .trim();

    if (!text) {
      return "Ollama finished but returned empty output.";
    }

    if (text.length > 1200) {
      text = text.slice(0, 1200) + "\n\n[truncated]";
    }

    return text;
  }

  function startAgent(taskText: string) {
    if (isRunning) {
      throw new Error("Agent is already running.");
    }

    isRunning = true;
    lastTask = taskText;
    writeFileSync(taskFile, taskText, "utf8");
    logLine(`Starting agent with task: ${taskText}`);

    const child = spawn(
      "powershell",
      [
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        runnerFile
      ],
      {
        cwd: repoRoot,
        stdio: ["ignore", "pipe", "pipe"]
      }
    );

    child.stdout.on("data", (chunk) => {
      logLine(`[stdout] ${String(chunk)}`);
    });

    child.stderr.on("data", (chunk) => {
      logLine(`[stderr] ${String(chunk)}`);
    });

    child.on("close", async () => {
  isRunning = false;

  const resultText = getResultText();

  try {
    await client.messages.create({
      from: whatsappFrom,
      to: whatsappTo,
      body: `Tsifulator run finished.\n\nTask:\n${lastTask}\n\nResult:\n${resultText}`
    });
  } catch (error) {
    logLine(`Failed to send WhatsApp completion message: ${String(error)}`);
  }
});

      try {
        if (ok) {
          await sendWhatsapp(
            `Tsifulator Ollama run finished.\n\nTask:\n${lastTask}\n\nResult:\n${resultText}`
          );
        } else {
          await sendWhatsapp(
            `Tsifulator Ollama run failed with exit code ${code}.\n\nTask:\n${lastTask}\n\nLast output:\n${resultText}`
          );
        }
      } catch (error) {
        logLine(`Failed to send WhatsApp completion message: ${String(error)}`);
      }
    });

    child.on("error", async (error) => {
      isRunning = false;
      logLine(`Agent process error: ${String(error)}`);

      try {
        await sendWhatsapp(`Tsifulator agent could not start.\n\nError:\n${String(error)}`);
      } catch (sendError) {
        logLine(`Failed to send WhatsApp startup error: ${String(sendError)}`);
      }
    });
  }

  app.get("/health", async () => {
    return {
      ok: true,
      running: isRunning,
      lastTask
    };
  });

  app.post("/whatsapp/incoming", async (request, reply) => {
    const body = request.body as Record<string, string | undefined>;
    const incomingText = (body.Body ?? "").trim();
    const from = body.From ?? "";

    const twiml = new twilio.twiml.MessagingResponse();

    if (normalizePhone(from) !== normalizePhone(whatsappTo)) {
      twiml.message("Unauthorized sender.");
      reply.type("text/xml").send(twiml.toString());
      return;
    }

    if (!incomingText) {
      twiml.message("Empty task. Reply with the coding task you want me to run.");
      reply.type("text/xml").send(twiml.toString());
      return;
    }

    if (isRunning) {
      twiml.message("Tsifulator agent is already running. Wait for the completion message, then send the next task.");
      reply.type("text/xml").send(twiml.toString());
      return;
    }

    try {
      startAgent(incomingText);
      twiml.message("Got it. Starting the Ollama coding run now.");
      reply.type("text/xml").send(twiml.toString());
      return;
    } catch (error) {
      twiml.message(`Could not start agent: ${String(error)}`);
      reply.type("text/xml").send(twiml.toString());
      return;
    }
  });

  await app.listen({ port, host: "0.0.0.0" });
  console.log(`WhatsApp agent listening on http://0.0.0.0:${port}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
