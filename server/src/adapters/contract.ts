import { AdapterContext, AdapterSuggestion } from "../types";

export interface AppAdapter {
  name: string;
  captureContext(context: AdapterContext): Record<string, unknown>;
  proposeActions(context: AdapterContext): AdapterSuggestion;
  validateAction(command: string): "safe" | "confirm" | "blocked";
  executeAction(): Promise<{ supported: boolean; output: string }>;
  emitEvents(context: AdapterContext): Array<{ type: string; payload: unknown }>;
}
