import { AppAdapter } from "./contract";
import { excelAdapter } from "./excel-adapter";
import { rstudioAdapter } from "./rstudio-adapter";
import { terminalAdapter } from "./terminal-adapter";

const adapters: Record<string, AppAdapter> = {
  terminal: terminalAdapter,
  excel: excelAdapter,
  rstudio: rstudioAdapter,
};

export function getAdapter(name: string): AppAdapter {
  return adapters[name] ?? terminalAdapter;
}

export function listAdapters(): string[] {
  return Object.keys(adapters);
}
