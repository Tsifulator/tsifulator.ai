@'
export interface ActionExecutionResult {
  supported: boolean;
  output: string;
}

export interface AdapterEvent {
  type: string;
  payload: unknown;
}
'@ | Set-Content .\server\src\shared-types.ts