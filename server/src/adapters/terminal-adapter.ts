import { classifyRisk } from "../risk";
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

export const terminalAdapter: AppAdapter = {
  name: "terminal",
  captureContext(context: AdapterContext) {
    return {
      cwd: context.cwd,
      lastOutput: context.lastOutput?.slice(0, 1000)
    };
  },
  proposeActions(context: AdapterContext) {
    return getTerminalSuggestion(context.message);
  },
  validateAction(command: string) {
    return classifyRisk(command);
  },
  async executeAction() {
    return {
      supported: false,
      output: "Execution handled by core policy executor"
    };
  },
  emitEvents(context: AdapterContext) {
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
};
