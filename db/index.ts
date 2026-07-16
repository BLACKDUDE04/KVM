import { drizzle } from "drizzle-orm/d1";
import * as schema from "./schema";
import { getEnv } from "../lib/env-context";

export function getDb() {
  return drizzle(getEnv().DB, { schema });
}
