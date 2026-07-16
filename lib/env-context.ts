import { AsyncLocalStorage } from "node:async_hooks";
import { createTursoDatabase } from "@/lib/turso-d1";

export type AppEnv = { DB: D1Database; ASSETS?: Fetcher; IMAGES?: unknown };
const key = Symbol.for("billflow.env.context");
const root = globalThis as typeof globalThis & { [key]?: AsyncLocalStorage<AppEnv> };
const storage = root[key] ??= new AsyncLocalStorage<AppEnv>();
let netlifyEnv: AppEnv | undefined;

export function runWithEnv<T>(env: AppEnv, fn: () => T | Promise<T>) {
  return storage.run(env, fn);
}

export function getEnv(): AppEnv {
  const env = storage.getStore();
  if (env?.DB) return env;

  const url = process.env.TURSO_DATABASE_URL?.trim();
  const authToken = process.env.TURSO_AUTH_TOKEN?.trim();
  if (!url) {
    throw new Error(
      "Database is not configured. Set TURSO_DATABASE_URL and TURSO_AUTH_TOKEN in Netlify, or provide the Cloudflare DB binding.",
    );
  }

  netlifyEnv ??= { DB: createTursoDatabase(url, authToken) };
  return netlifyEnv;
}
