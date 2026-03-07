# Adapter Interfaces

## Overview

Tsifulator.ai uses an **adapter pattern** to support multiple client environments (terminal, Excel, RStudio, etc.) through a common interface. Each adapter customizes context capture, action proposals, validation, execution, and telemetry for its target environment.

## Interface: `AppAdapter`

Defined in `server/src/adapters/contract.ts`:

```typescript
interface AppAdapter {
  name: string;
  captureContext(context: AdapterContext): Record<string, unknown>;
  proposeActions(context: AdapterContext): AdapterSuggestion;
  validateAction(command: string): "safe" | "confirm" | "blocked";
  executeAction(): Promise<{ supported: boolean; output: string }>;
  emitEvents(context: AdapterContext): Array<{ type: string; payload: unknown }>;
}
```

### Methods

| Method | Purpose | Returns |
|---|---|---|
| `captureContext` | Extract environment-specific context from the request | Key-value metadata |
| `proposeActions` | Generate action suggestions based on user message | Text response + optional command |
| `validateAction` | Classify a command's risk level for this adapter | `"safe"` / `"confirm"` / `"blocked"` |
| `executeAction` | Adapter-specific execution (if supported) | `{ supported, output }` |
| `emitEvents` | Generate telemetry events for this adapter | Array of `{ type, payload }` |

### Supporting types

```typescript
interface AdapterContext {
  userId: string;
  sessionId: string;
  message: string;
  cwd?: string;
  lastOutput?: string;
}

interface AdapterSuggestion {
  text: string;
  command?: string;
}
```

## Current adapters

| Adapter | Status | File |
|---|---|---|
| `terminal` | ✅ Complete | `server/src/adapters/terminal-adapter.ts` |
| `excel` | 🏗 Scaffold | `server/src/adapters/excel-adapter.ts` |
| `rstudio` | ❌ Not started | — |

## Adding a new adapter

1. Create `server/src/adapters/<name>-adapter.ts`
2. Implement the `AppAdapter` interface
3. Register in `server/src/adapters/registry.ts`:
   ```typescript
   import { newAdapter } from "./<name>-adapter";
   const adapters: Record<string, AppAdapter> = {
     terminal: terminalAdapter,
     excel: excelAdapter,
     "<name>": newAdapter,
   };
   ```
4. Add tests in `tests/<name>-adapter.test.mjs`

## Registry

`server/src/adapters/registry.ts` provides:
- `getAdapter(name)` — returns adapter by name (falls back to terminal)
- `listAdapters()` — returns registered adapter names (shown in `/health` response)
