import fs from "node:fs";
import path from "node:path";
import { z } from "zod";

const isProd = process.env.NODE_ENV === "production";

const envSchema = z.object({
  NODE_ENV: z.enum(["development", "production", "test"]).default("development"),
  PORT: z.coerce.number().default(4000),
  HOST: z.string().default("0.0.0.0"),
  DB_PATH: z.string().default("./data/tsifulator.db"),
  JWT_DEV_SECRET: z.string().default("change_me_in_beta"),
  OPENAI_API_KEY: z.string().optional(),
  CORS_ORIGIN: z.string().default(isProd ? "" : "*"),
  LOG_LEVEL: z.enum(["fatal", "error", "warn", "info", "debug", "trace"]).default(isProd ? "info" : "debug"),
});

export type AppConfig = z.infer<typeof envSchema>;

export function getConfig(): AppConfig {
  const parsed = envSchema.parse(process.env);

  // In production, enforce real secrets
  if (parsed.NODE_ENV === "production") {
    if (parsed.JWT_DEV_SECRET === "change_me_in_beta") {
      throw new Error("JWT_DEV_SECRET must be changed from default in production");
    }
  }

  const resolvedDbPath = path.resolve(parsed.DB_PATH);
  const dbDir = path.dirname(resolvedDbPath);
  if (!fs.existsSync(dbDir)) {
    fs.mkdirSync(dbDir, { recursive: true });
  }

  return {
    ...parsed,
    DB_PATH: resolvedDbPath
  };
}
