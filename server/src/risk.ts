import { ActionRisk } from "./shared-types";

const BLOCKED_PATTERNS = [
  /rm\s+-rf/i,
  /del\s+\/s\s+\/q/i,
  /remove-item\b.*-recurse/i,
  /remove-item\b.*-force/i,
  /rd\s+\/s\s+\/q/i,
  /\*\.\*/,
  /sudo\s+/i,
  /runas\s+/i,
  /chmod\s+-R\s+777/i,
  /chown\s+-R/i
];

const CONFIRM_PATTERNS = [/npm\s+install/i, /git\s+(reset|clean)/i, /docker\s+/i];

export type RiskLevel = ActionRisk;

export function classifyRisk(command: string): RiskLevel {
  if (BLOCKED_PATTERNS.some((pattern) => pattern.test(command))) {
    return "blocked";
  }

  if (CONFIRM_PATTERNS.some((pattern) => pattern.test(command))) {
    return "confirm";
  }

  return "safe";
}

export function redactSecrets(output: string): string {
  return output
    .replace(/(api[_-]?key\s*[=:]\s*)([^\s]+)/gi, "$1[REDACTED]")
    .replace(/(token\s*[=:]\s*)([^\s]+)/gi, "$1[REDACTED]")
    .replace(/(password\s*[=:]\s*)([^\s]+)/gi, "$1[REDACTED]");
}

export function boundOutput(output: string, maxChars = 4000): string {
  if (output.length <= maxChars) {
    return output;
  }

  return `${output.slice(0, maxChars)}\n...[truncated]`;
}

