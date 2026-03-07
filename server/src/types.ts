export interface ChatRequest {
  userId: string;
  sessionId?: string;
  message: string;
  cwd?: string;
  lastOutput?: string;
}

export interface ChatResponse {
  sessionId: string;
  text: string;
  proposal: {
    id: string;
    sessionId: string;
    command: string;
    risk: "safe" | "confirm" | "blocked";
  } | null;
}

export interface AdapterContext {
  userId: string;
  sessionId: string;
  message: string;
  cwd?: string;
  lastOutput?: string;
}

export interface AdapterSuggestion {
  text: string;
  command?: string;
}

export interface AuthUser {
  id: string;
  email: string;
}
