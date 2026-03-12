export type ActionRisk = "safe" | "confirm" | "blocked";

export interface ActionExecutionResult {
  supported: boolean;
  output: string;
}

export interface AdapterEvent {
  type: string;
  payload: unknown;
}
