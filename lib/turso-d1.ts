import {
  createClient,
  type Client,
  type InStatement,
  type InValue,
  type ResultSet,
} from "@tursodatabase/serverless/compat";

type D1Meta = {
  duration: number;
  size_after: number;
  rows_read: number;
  rows_written: number;
  last_row_id: number;
  changed_db: boolean;
  changes: number;
};

function normalizeValue(value: unknown): InValue {
  if (value === undefined) return null;
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "bigint" ||
    typeof value === "boolean" ||
    value instanceof ArrayBuffer ||
    value instanceof Uint8Array ||
    value instanceof Date
  ) {
    return value;
  }
  return String(value);
}

function rowsFrom(result: ResultSet) {
  return result.rows.map((row) =>
    Object.fromEntries(
      result.columns.map((column, index) => [column, row[index]]),
    ),
  );
}

function resultFrom(result: ResultSet, startedAt: number) {
  const changes = Number(result.rowsAffected || 0);
  const meta: D1Meta = {
    duration: Math.max(0, performance.now() - startedAt),
    size_after: 0,
    rows_read: result.rows.length,
    rows_written: changes,
    last_row_id: Number(result.lastInsertRowid || 0),
    changed_db: changes > 0,
    changes,
  };

  return {
    success: true,
    results: rowsFrom(result),
    meta,
  };
}

class TursoPreparedStatement {
  readonly sql: string;
  readonly args: InValue[];
  private readonly client: Client;

  constructor(client: Client, sql: string, args: InValue[] = []) {
    this.client = client;
    this.sql = sql;
    this.args = args;
  }

  bind(...values: unknown[]) {
    return new TursoPreparedStatement(
      this.client,
      this.sql,
      values.map(normalizeValue),
    );
  }

  async all<T = Record<string, unknown>>() {
    const startedAt = performance.now();
    const result = await this.client.execute(this.statement());
    return resultFrom(result, startedAt) as unknown as D1Result<T>;
  }

  async first<T = Record<string, unknown>>(column?: string) {
    const result = await this.all<Record<string, unknown>>();
    const row = result.results[0];
    if (!row) return null;
    return (column ? row[column] : row) as T;
  }

  async run<T = Record<string, unknown>>() {
    return this.all<T>();
  }

  async raw<T = unknown[]>(options?: { columnNames?: boolean }) {
    const result = await this.client.execute(this.statement());
    const values = result.rows.map((row) =>
      result.columns.map((_, index) => row[index]),
    );
    return (options?.columnNames
      ? [result.columns, ...values]
      : values) as unknown as T[];
  }

  statement(): InStatement {
    return { sql: this.sql, args: this.args };
  }
}

class TursoD1Database {
  private readonly client: Client;

  constructor(url: string, authToken?: string) {
    this.client = createClient({ url, authToken });
  }

  prepare(sql: string) {
    return new TursoPreparedStatement(this.client, sql);
  }

  async batch<T = Record<string, unknown>>(
    statements: TursoPreparedStatement[],
  ) {
    if (!statements.length) return [];
    const startedAt = performance.now();
    const results = await this.client.batch(
      statements.map((statement) => statement.statement()),
      "write",
    );
    return results.map((result) =>
      resultFrom(result, startedAt),
    ) as unknown as D1Result<T>[];
  }
}

export function createTursoDatabase(url: string, authToken?: string) {
  return new TursoD1Database(url, authToken) as unknown as D1Database;
}
