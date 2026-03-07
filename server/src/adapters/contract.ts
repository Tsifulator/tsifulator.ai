import { AdapterContext, AdapterSuggestion } from "../types";
import { SharedMemory } from "../shared-memory";

export interface AppAdapter {
  name: string;
  captureContext(context: AdapterContext): Record<string, unknown>;
  proposeActions(context: AdapterContext): AdapterSuggestion;
  validateAction(command: string): "safe" | "confirm" | "blocked";
  executeAction(context: AdapterContext, memory?: SharedMemory): Promise<{ supported: boolean; output: string }>;
  emitEvents(context: AdapterContext): Array<{ type: string; payload: unknown }>;

  /** Hook called after chat to persist adapter-specific memory entries. */
  saveToMemory?(context: AdapterContext, memory: SharedMemory): void;
}
