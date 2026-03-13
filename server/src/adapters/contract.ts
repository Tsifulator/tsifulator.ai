import { AdapterContext, AdapterSuggestion } from "../types";
import { SharedMemory } from "../shared-memory";
import { ActionExecutionResult, AdapterEvent, ActionRisk } from "../shared-types";

export interface AppAdapter {
  name: string;
  captureContext(context: AdapterContext): Record<string, unknown>;
  proposeActions(context: AdapterContext): AdapterSuggestion;
  validateAction(command: string): ActionRisk;
  executeAction(context: AdapterContext, memory?: SharedMemory): Promise<ActionExecutionResult>;
  emitEvents(context: AdapterContext): AdapterEvent[];

  /** Hook called after chat to persist adapter-specific memory entries. */
  saveToMemory?(context: AdapterContext, memory: SharedMemory): void;
}

