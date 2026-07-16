import { getEnv } from "../../../lib/env-context";

const resources: Record<
  string,
  { table: string; fields: string[]; required: string[] }
> = {
  customers: {
    table: "customers",
    fields: [
      "name",
      "owner_name",
      "code",
      "gstin",
      "phone",
      "registration_type",
      "state",
      "email",
      "address",
      "balance",
    ],
    required: ["name"],
  },
  suppliers: {
    table: "suppliers",
    fields: ["name", "code", "phone", "state", "email", "gstin", "address"],
    required: ["name"],
  },
  products: {
    table: "products",
    fields: [
      "name",
      "sku",
      "hsn_code",
      "category",
      "stock",
      "reorder_level",
      "purchase_rate",
      "price",
      "gst_rate",
      "created_at",
    ],
    required: ["name", "sku"],
  },
  invoices: {
    table: "invoices",
    fields: [
      "invoice_no",
      "customer_id",
      "customer_name",
      "customer_code",
      "kind",
      "items_json",
      "subtotal",
      "tax",
      "amount",
      "status",
      "created_at",
    ],
    required: ["invoice_no", "customer_name", "amount", "created_at"],
  },
  purchases: {
    table: "purchases",
    fields: [
      "bill_no",
      "supplier",
      "supplier_code",
      "gstin",
      "item",
      "item_code",
      "quantity",
      "rate",
      "gst_amount",
      "total",
      "purchase_date",
      "received_date",
      "import_batch",
    ],
    required: [
      "bill_no",
      "supplier",
      "item",
      "quantity",
      "rate",
      "total",
      "purchase_date",
    ],
  },
  returns: {
    table: "sales_returns",
    fields: [
      "credit_note_no",
      "invoice_id",
      "customer_id",
      "amount",
      "reason",
      "created_at",
    ],
    required: [
      "credit_note_no",
      "invoice_id",
      "customer_id",
      "amount",
      "created_at",
    ],
  },
  payments: {
    table: "payments",
    fields: [
      "customer_id",
      "invoice_id",
      "amount",
      "method",
      "reference",
      "collected_at",
    ],
    required: ["customer_id", "amount", "method", "collected_at"],
  },
  accounts: {
    table: "accounts",
    fields: [
      "name",
      "type",
      "bank_name",
      "account_last4",
      "opening_balance",
      "active",
    ],
    required: ["name", "type"],
  },
  transactions: {
    table: "account_transactions",
    fields: [
      "account_id",
      "direction",
      "amount",
      "particulars",
      "reference",
      "transaction_date",
    ],
    required: [
      "account_id",
      "direction",
      "amount",
      "particulars",
      "transaction_date",
    ],
  },
  users: {
    table: "users",
    fields: ["name", "email", "role", "enabled"],
    required: ["name", "email"],
  },
  settings: {
    table: "settings",
    fields: ["business", "gstin", "prefix", "gst"],
    required: ["business"],
  },
};
let ready = false;
let schemaPromise: Promise<void> | null = null;
export async function ensureSchema() {
  if (ready) return;
  if (!schemaPromise) schemaPromise = prepareSchema();
  try {
    await schemaPromise;
    ready = true;
  } catch (error) {
    schemaPromise = null;
    throw error;
  }
}
async function prepareSchema() {
  const db = getEnv().DB;
  await db.batch([
    db.prepare(
      "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, email TEXT NOT NULL UNIQUE, role TEXT NOT NULL DEFAULT 'staff', enabled INTEGER NOT NULL DEFAULT 1)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS customers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, gstin TEXT, phone TEXT, balance REAL NOT NULL DEFAULT 0)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS products (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, sku TEXT NOT NULL UNIQUE, category TEXT NOT NULL DEFAULT 'General', stock REAL NOT NULL DEFAULT 0, reorder_level REAL NOT NULL DEFAULT 5, purchase_rate REAL NOT NULL DEFAULT 0, price REAL NOT NULL DEFAULT 0, gst_rate REAL NOT NULL DEFAULT 18, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS invoices (id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_no TEXT NOT NULL UNIQUE, customer_id INTEGER, customer_name TEXT NOT NULL, kind TEXT NOT NULL DEFAULT 'invoice', items_json TEXT NOT NULL DEFAULT '[]', subtotal REAL NOT NULL DEFAULT 0, tax REAL NOT NULL DEFAULT 0, amount REAL NOT NULL, status TEXT NOT NULL DEFAULT 'pending', created_at TEXT NOT NULL)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS purchases (id INTEGER PRIMARY KEY AUTOINCREMENT, bill_no TEXT NOT NULL, supplier TEXT NOT NULL, gstin TEXT, item TEXT NOT NULL, quantity REAL NOT NULL, rate REAL NOT NULL, gst_amount REAL NOT NULL DEFAULT 0, total REAL NOT NULL, purchase_date TEXT NOT NULL, import_batch TEXT)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS sales_returns (id INTEGER PRIMARY KEY AUTOINCREMENT, credit_note_no TEXT NOT NULL UNIQUE, invoice_id INTEGER NOT NULL, customer_id INTEGER NOT NULL, amount REAL NOT NULL, reason TEXT, created_at TEXT NOT NULL)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS payments (id INTEGER PRIMARY KEY AUTOINCREMENT, customer_id INTEGER NOT NULL, invoice_id INTEGER, amount REAL NOT NULL, method TEXT NOT NULL, reference TEXT, collected_at TEXT NOT NULL)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS accounts (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, type TEXT NOT NULL, bank_name TEXT, account_last4 TEXT, opening_balance REAL NOT NULL DEFAULT 0, active INTEGER NOT NULL DEFAULT 1)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS account_transactions (id INTEGER PRIMARY KEY AUTOINCREMENT, account_id INTEGER NOT NULL, direction TEXT NOT NULL, amount REAL NOT NULL, particulars TEXT NOT NULL, reference TEXT, transaction_date TEXT NOT NULL)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS settings (id INTEGER PRIMARY KEY AUTOINCREMENT, business TEXT NOT NULL, gstin TEXT, prefix TEXT NOT NULL DEFAULT 'INV-', gst TEXT NOT NULL DEFAULT '18')",
    ),
  ]);
  const addColumn = async (table: string, name: string, type: string) => {
    try {
      await db.prepare(`ALTER TABLE ${table} ADD COLUMN ${name} ${type}`).run();
    } catch (error) {
      if (!String(error).toLowerCase().includes("duplicate column name"))
        throw error;
    }
  };
  const migrations: Record<string, string> = {
    owner_name: "TEXT",
    code: "TEXT",
    registration_type: "TEXT",
    state: "TEXT",
    email: "TEXT",
    address: "TEXT",
  };
  const customerCols = (
    await db.prepare("PRAGMA table_info(customers)").all()
  ).results.map((x) => String((x as Record<string, unknown>).name));
  for (const [name, type] of Object.entries(migrations))
    if (!customerCols.includes(name)) await addColumn("customers", name, type);
  const productCols = (
    await db.prepare("PRAGMA table_info(products)").all()
  ).results.map((x) => String((x as Record<string, unknown>).name));
  for (const [name, type] of Object.entries({
    hsn_code: "TEXT",
    purchase_rate: "REAL NOT NULL DEFAULT 0",
    created_at: "TEXT",
  }))
    if (!productCols.includes(name)) await addColumn("products", name, type);
  await db
    .prepare(
      "UPDATE products SET created_at = datetime('now') WHERE created_at IS NULL OR created_at = ''",
    )
    .run();
  const invoiceCols = (
    await db.prepare("PRAGMA table_info(invoices)").all()
  ).results.map((x) => String((x as Record<string, unknown>).name));
  if (!invoiceCols.includes("customer_code"))
    await addColumn("invoices", "customer_code", "TEXT");
  const purchaseCols = (
    await db.prepare("PRAGMA table_info(purchases)").all()
  ).results.map((x) => String((x as Record<string, unknown>).name));
  for (const name of ["supplier_code", "item_code", "received_date"])
    if (!purchaseCols.includes(name))
      await addColumn("purchases", name, "TEXT");
  await db
    .prepare(
      "CREATE TABLE IF NOT EXISTS suppliers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, code TEXT, phone TEXT, state TEXT, email TEXT, gstin TEXT, address TEXT)",
    )
    .run();
}

function cfg(request: Request) {
  const key = new URL(request.url).searchParams.get("resource") || "";
  const value = resources[key];
  if (!value) throw new Error("Unknown resource");
  return { key, ...value };
}
export async function authorize(request: Request) {
  const email =
    request.headers.get("oai-authenticated-user-email") ||
    (new URL(request.url).hostname === "terminal.local" ? "owner@local" : null);
  if (!email) return false;
  const db = getEnv().DB;
  const found = await db
    .prepare("SELECT enabled FROM users WHERE email = ? LIMIT 1")
    .bind(email)
    .first<{ enabled: number }>();
  if (!found) {
    const count = await db
      .prepare("SELECT COUNT(*) AS count FROM users")
      .first<{ count: number }>();
    if ((count?.count || 0) === 0)
      await db
        .prepare(
          "INSERT OR IGNORE INTO users(name,email,role,enabled) VALUES(?,?,?,1)",
        )
        .bind(email.split("@")[0], email, "Admin")
        .run();
    return true;
  }
  return Boolean(found.enabled);
}
function clean(body: Record<string, unknown>, fields: string[]) {
  return Object.fromEntries(
    fields.filter((f) => body[f] !== undefined).map((f) => [f, body[f]]),
  );
}
export async function GET(request: Request) {
  try {
    await ensureSchema();
    if (!(await authorize(request)))
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 403 },
      );
    const c = cfg(request);
    const rows = await getEnv()
      .DB.prepare(`SELECT * FROM ${c.table} ORDER BY id DESC LIMIT 10000`)
      .all();
    return Response.json({ rows: rows.results });
  } catch (e) {
    return Response.json(
      { error: e instanceof Error ? e.message : "Unable to load data" },
      { status: 500 },
    );
  }
}
export async function POST(request: Request) {
  try {
    await ensureSchema();
    if (!(await authorize(request)))
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 403 },
      );
    const c = cfg(request),
      payload = (await request.json()) as
        | Record<string, unknown>
        | Record<string, unknown>[];
    const stampProduct = (row: Record<string, unknown>) => {
      if (c.key === "products" && !row.created_at)
        row.created_at = new Date().toISOString();
    };
    if (Array.isArray(payload)) {
      const db = getEnv().DB;
      payload.forEach(stampProduct);
      if (c.key === "invoices") {
        const customers = (
          await db.prepare("SELECT id,name,code FROM customers").all()
        ).results as Record<string, unknown>[];
        const missing = new Map<string, { name: string; code: string }>();
        for (const row of payload) {
          const name = String(row.customer_name || "Unknown Customer"),
            code = String(row.customer_code || "");
          if (
            !customers.some(
              (x) =>
                (code && String(x.code) === code) ||
                String(x.name).toLowerCase() === name.toLowerCase(),
            )
          )
            missing.set(code || name.toLowerCase(), { name, code });
        }
        if (missing.size)
          await db.batch(
            [...missing.values()].map((x) =>
              db
                .prepare(
                  "INSERT INTO customers(name,code,balance) VALUES(?,?,0)",
                )
                .bind(x.name, x.code),
            ),
          );
        const allCustomers = (
          await db.prepare("SELECT id,name,code FROM customers").all()
        ).results as Record<string, unknown>[];
        for (const row of payload) {
          const name = String(row.customer_name || "Unknown Customer"),
            code = String(row.customer_code || "");
          row.customer_id = allCustomers.find(
            (x) =>
              (code && String(x.code) === code) ||
              String(x.name).toLowerCase() === name.toLowerCase(),
          )?.id;
        }
      }
      const statements = payload.map((body) => {
        for (const f of c.required)
          if (body[f] === undefined || body[f] === "")
            throw new Error(`${f} is required`);
        const values = clean(body, c.fields),
          keys = Object.keys(values);
        return db
          .prepare(
            `INSERT OR IGNORE INTO ${c.table} (${keys.join(",")}) VALUES (${keys.map(() => "?").join(",")})`,
          )
          .bind(...keys.map((k) => values[k]));
      });
      if (statements.length) await db.batch(statements);
      return Response.json({ count: statements.length }, { status: 201 });
    }
    const body = payload;
    stampProduct(body);
    if (c.key === "invoices" && !Number(body.customer_id)) {
      const db = getEnv().DB,
        customerName = String(body.customer_name || "Unknown Customer"),
        customerCode = String(body.customer_code || ""),
        existing = await db
          .prepare(
            "SELECT id FROM customers WHERE (code != '' AND code = ?) OR lower(name) = lower(?) LIMIT 1",
          )
          .bind(customerCode, customerName)
          .first<{ id: number }>();
      if (existing?.id) body.customer_id = existing.id;
      else {
        const created = await db
          .prepare(
            "INSERT INTO customers(name, code, balance) VALUES(?, ?, 0) RETURNING id",
          )
          .bind(customerName, customerCode)
          .first<{ id: number }>();
        body.customer_id = created?.id;
      }
    }
    for (const f of c.required)
      if (body[f] === undefined || body[f] === "")
        return Response.json({ error: `${f} is required` }, { status: 400 });
    const values = clean(body, c.fields),
      keys = Object.keys(values);
    const result = await getEnv()
      .DB.prepare(
        `INSERT INTO ${c.table} (${keys.join(",")}) VALUES (${keys.map(() => "?").join(",")}) RETURNING *`,
      )
      .bind(...keys.map((k) => values[k]))
      .first();
    return Response.json({ row: result }, { status: 201 });
  } catch (e) {
    return Response.json(
      { error: e instanceof Error ? e.message : "Unable to save data" },
      { status: 500 },
    );
  }
}
export async function PATCH(request: Request) {
  try {
    await ensureSchema();
    if (!(await authorize(request)))
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 403 },
      );
    const c = cfg(request),
      payload = (await request.json()) as
        | Record<string, unknown>
        | Record<string, unknown>[];
    if (Array.isArray(payload)) {
      const db = getEnv().DB,
        statements = payload.map((body) => {
          const id = Number(body.id);
          if (!id) throw new Error("id is required");
          const values = clean(body, c.fields),
            keys = Object.keys(values);
          return db
            .prepare(
              `UPDATE ${c.table} SET ${keys.map((k) => `${k}=?`).join(",")} WHERE id=?`,
            )
            .bind(...keys.map((k) => values[k]), id);
        });
      if (statements.length) await db.batch(statements);
      return Response.json({ count: statements.length });
    }
    const body = payload;
    const id = Number(body.id);
    if (!id) return Response.json({ error: "id is required" }, { status: 400 });
    const values = clean(body, c.fields),
      keys = Object.keys(values);
    if (!keys.length)
      return Response.json({ error: "No fields to update" }, { status: 400 });
    const row = await getEnv()
      .DB.prepare(
        `UPDATE ${c.table} SET ${keys.map((k) => `${k}=?`).join(",")} WHERE id=? RETURNING *`,
      )
      .bind(...keys.map((k) => values[k]), id)
      .first();
    return Response.json({ row });
  } catch (e) {
    return Response.json(
      { error: e instanceof Error ? e.message : "Unable to update data" },
      { status: 500 },
    );
  }
}
export async function DELETE(request: Request) {
  try {
    await ensureSchema();
    if (!(await authorize(request)))
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 403 },
      );
    const c = cfg(request),
      url = new URL(request.url),
      id = Number(url.searchParams.get("id"));
    if (url.searchParams.get("all") === "1") {
      if (c.key === "invoices" && url.searchParams.get("kind")) {
        await getEnv()
          .DB.prepare("DELETE FROM invoices WHERE kind = ?")
          .bind(url.searchParams.get("kind"))
          .run();
        return Response.json({ ok: true });
      }
      if (c.key === "transactions" && url.searchParams.get("type")) {
        await getEnv()
          .DB.prepare(
            "DELETE FROM account_transactions WHERE account_id IN (SELECT id FROM accounts WHERE lower(type) = ?)",
          )
          .bind(url.searchParams.get("type"))
          .run();
        return Response.json({ ok: true });
      }
      if (c.key === "accounts" && url.searchParams.get("type")) {
        await getEnv()
          .DB.prepare("DELETE FROM accounts WHERE lower(type) = ?")
          .bind(url.searchParams.get("type"))
          .run();
        return Response.json({ ok: true });
      }
      await getEnv().DB.prepare(`DELETE FROM ${c.table}`).run();
      return Response.json({ ok: true });
    }
    if (!id) return Response.json({ error: "id is required" }, { status: 400 });
    await getEnv()
      .DB.prepare(`DELETE FROM ${c.table} WHERE id=?`)
      .bind(id)
      .run();
    return Response.json({ ok: true });
  } catch (e) {
    return Response.json(
      { error: e instanceof Error ? e.message : "Unable to delete data" },
      { status: 500 },
    );
  }
}
