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
    if (!token || !token.startsWith("dev-")) {
      return reply.code(401).send({ error: "Missing or invalid bearer token" });
    }

    const userId = token.slice(4);
    const user = db.getUserById(userId);
    if (!user) {
      return reply.code(401).send({ error: "Token user not found" });
    }

    (request as FastifyRequest & { authUser: AuthUser }).authUser = user;
  };
}
