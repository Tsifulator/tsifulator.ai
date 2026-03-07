import { FastifyReply, FastifyRequest } from "fastify";
import { AppDb } from "./db";
import { AuthUser } from "./types";

function extractBearer(authHeader: string | undefined): string | null {
  if (!authHeader) {
    return null;
  }

  const [scheme, token] = authHeader.split(" ");
  if (scheme?.toLowerCase() !== "bearer" || !token) {
    return null;
  }

  return token;
}

export function requireDevAuth(db: AppDb) {
  return async function authPreHandler(request: FastifyRequest, reply: FastifyReply) {
    const token = extractBearer(request.headers.authorization);
    if (!token) {
      return reply.code(401).send({
        error: "Authentication required",
        hint: "Include an Authorization header: Bearer <token>. Get a token via POST /auth/dev-login.",
      });
    }

    // API key authentication (tsk_...)
    if (token.startsWith("tsk_")) {
      const user = db.validateApiKey(token);
      if (!user) {
        return reply.code(401).send({
          error: "Invalid or revoked API key",
          hint: "Check that your API key is correct and has not been revoked.",
        });
      }
      (request as FastifyRequest & { authUser: AuthUser }).authUser = user;
      return;
    }

    // Dev-login authentication (dev-...)
    if (!token.startsWith("dev-")) {
      return reply.code(401).send({
        error: "Invalid token format",
        hint: "Use a dev-login token (dev-...) or an API key (tsk_...). Get a token via POST /auth/dev-login.",
      });
    }

    const userId = token.slice(4);
    const user = db.getUserById(userId);
    if (!user) {
      return reply.code(401).send({
        error: "Token references unknown user",
        hint: "This token may be stale. Re-authenticate via POST /auth/dev-login.",
      });
    }

    (request as FastifyRequest & { authUser: AuthUser }).authUser = user;
  };
}
