import { AsyncLocalStorage } from "node:async_hooks";
import { createTursoDatabase } from "@/lib/turso-d1";

export type AppEnv = { DB: D1Database; ASSETS?: Fetcher; IMAGES?: unknown };
const key = Symbol.for("billflow.env.context");
const root = globalThis as typeof globalThis & { [key]?: AsyncLocalStorage<AppEnv> };
const storage = root[key] ??= new AsyncLocalStorage<AppEnv>();
let netlifyEnv: AppEnv | undefined;

function cleanEnvironmentValue(value: string | undefined) {
  const trimmed = value?.trim() || "";
  if (
    trimmed.length >= 2 &&
    ((trimmed.startsWith('"') && trimmed.endsWith('"')) ||
      (trimmed.startsWith("'") && trimmed.endsWith("'")))
  )
    return trimmed.slice(1, -1).trim();
  return trimmed;
}

function databaseCredentials() {
  const rawUrl = cleanEnvironmentValue(process.env.TURSO_DATABASE_URL),
    authToken = cleanEnvironmentValue(process.env.TURSO_AUTH_TOKEN);
  if (!rawUrl)
    throw new Error(
      "Database is not configured. Add TURSO_DATABASE_URL and TURSO_AUTH_TOKEN to the Netlify Functions environment.",
    );
  if (!/^libsql:\/\/|^https:\/\//i.test(rawUrl))
    throw new Error(
      "TURSO_DATABASE_URL is invalid. Copy the database URL from Turso; it must start with libsql:// or https://.",
    );

  const url = rawUrl.replace(/\/+$/, "");
  if (/\.turso\.io$/i.test(url) && !authToken)
    throw new Error(
      "TURSO_AUTH_TOKEN is missing. Create a token for the same Turso database and add it to the Netlify Functions environment.",
    );
  return { url, authToken: authToken || undefined };
}

export function runWithEnv<T>(env: AppEnv, fn: () => T | Promise<T>) {
  return storage.run(env, fn);
}

export function getEnv(): AppEnv {
  const env = storage.getStore();
  if (env?.DB) return env;

  const { url, authToken } = databaseCredentials();
  netlifyEnv ??= { DB: createTursoDatabase(url, authToken) };
  return netlifyEnv;
}
