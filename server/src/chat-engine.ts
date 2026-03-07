import { AppDb } from "./db";
import { boundOutput, redactSecrets } from "./risk";
import { ChatRequest, ChatResponse } from "./types";
import { getAdapter } from "./adapters/registry";

export function handleChat(db: AppDb, request: ChatRequest): ChatResponse {
  const sessionId = request.sessionId ?? db.createSession(request.userId).id;
  const adapter = getAdapter("terminal");

  db.saveMessage(sessionId, "user", request.message);

  const context = {
    userId: request.userId,
    sessionId,
    message: request.message,
    cwd: request.cwd,
    lastOutput: request.lastOutput
  };

  const captured = adapter.captureContext(context);
  db.logEvent(sessionId, "chat_user_message", {
    message: request.message,
    context: {
      ...captured,
      lastOutput: boundOutput(request.lastOutput ?? "", 1000)
    }
  });

  for (const event of adapter.emitEvents(context)) {
    db.logEvent(sessionId, event.type, event.payload);
  }

  const assistant = adapter.proposeActions(context);
  db.saveMessage(sessionId, "assistant", assistant.text);
  db.logEvent(sessionId, "chat_assistant_message", { text: assistant.text });

  let proposal: ChatResponse["proposal"] = null;
  if (assistant.command) {
    const risk = adapter.validateAction(assistant.command);
    proposal = db.saveActionProposal(sessionId, assistant.command, risk);
    db.logEvent(sessionId, "action_proposed", proposal);
  }

  db.saveAdapterState("terminal", redactSecrets(JSON.stringify(captured)));

  return {
    sessionId,
    text: assistant.text,
    proposal
  };
}
