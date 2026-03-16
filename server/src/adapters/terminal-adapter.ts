import { classifyRisk } from "../risk";
import { SharedMemory } from "../shared-memory";
import { AdapterContext, AdapterSuggestion } from "../types";
import { AppAdapter } from "./contract";

function getTerminalSuggestion(message: string): AdapterSuggestion {
  const lower = message.toLowerCase();

  if (lower.startsWith("cmd:")) {
    return {
      text: "I drafted a command proposal from your explicit cmd request.",
      command: message.slice(4).trim()
    };
  }

  if (lower.includes("list files") || lower.includes("show files")) {
    return {
      text: "I can list files in the current directory after your approval.",
      command: "Get-ChildItem"
    };
  }

  if (lower.includes("run test") || lower.includes("run tests")) {
    return {
      text: "I can run tests once you approve execution.",
      command: "npm test"
    };
  }

  return {
    text: "Received. I stored your message and can propose actions if you prefix with 'cmd:' or ask to list files/run tests."
  };
}

function captureTerminalContext(context: AdapterContext): Record<string, unknown> {
  return {
    cwd: context.cwd,
    lastOutput: context.lastOutput?.slice(0, 1000)
  };
}

function emitTerminalEvents(context: AdapterContext) {
  return [
    {
      type: "terminal_context_captured",
      payload: {
        cwd: context.cwd,
        hasLastOutput: Boolean(context.lastOutput)
      }
    }
  ];
}

function saveTerminalMemory(context: AdapterContext, memory: SharedMemory): void {
  if (context.cwd) {
    memory.set(context.userId, "terminal", "last_cwd", context.cwd, context.sessionId);
  }
  memory.set(context.userId, "terminal", "last_prompt", context.message.slice(0, 500), context.sessionId);
}

export const terminalAdapter: AppAdapter = {
  name: "terminal",
  captureContext: captureTerminalContext,
  proposeActions(context: AdapterContext) {
    return getTerminalSuggestion(context.message);
  },
  validateAction: classifyRisk,
  async executeAction() {
    return {
      supported: false,
      output: "Execution handled by core policy executor"
    };
  },
  emitEvents: emitTerminalEvents,
  saveToMemory: saveTerminalMemory
};



