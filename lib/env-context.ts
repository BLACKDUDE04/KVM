import { AsyncLocalStorage } from "node:async_hooks";

export type AppEnv = { DB: D1Database; ASSETS?: Fetcher; IMAGES?: unknown };
const key = Symbol.for("billflow.env.context");
const root = globalThis as typeof globalThis & { [key]?: AsyncLocalStorage<AppEnv> };
const storage = root[key] ??= new AsyncLocalStorage<AppEnv>();

export function runWithEnv<T>(env: AppEnv, fn: () => T | Promise<T>) {
  return storage.run(env, fn);
}

export function getEnv(): AppEnv {
  const env = storage.getStore();
  if (!env?.DB) throw new Error("Database binding is unavailable");
  return env;
}
