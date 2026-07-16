import { getEnv } from "../../../lib/env-context";
import {
  createPassword,
  getSessionUser,
  validateUserId,
} from "../../../lib/auth";

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
      "items_json",
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
      "account_id",
      "transaction_id",
      "amount",
      "method",
      "reference",
      "collected_at",
    ],
    required: ["customer_id", "account_id", "amount", "method", "collected_at"],
  },
  expenses: {
    table: "expenses",
    fields: [
      "expense_date",
      "category",
      "payee",
      "description",
      "reference",
      "payment_method",
      "account_id",
      "transaction_id",
      "amount",
      "created_at",
    ],
    required: [
      "expense_date",
      "category",
      "description",
      "payment_method",
      "account_id",
      "amount",
      "created_at",
    ],
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
    fields: [
      "name",
      "user_id",
      "email",
      "role",
      "enabled",
      "phone",
      "designation",
      "company",
      "timezone",
    ],
    required: ["name", "user_id"],
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
      "CREATE TABLE IF NOT EXISTS auth_sessions (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, token_hash TEXT NOT NULL UNIQUE, expires_at TEXT NOT NULL, created_at TEXT NOT NULL)",
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
    db.prepare(
      "CREATE TABLE IF NOT EXISTS expenses (id INTEGER PRIMARY KEY AUTOINCREMENT, expense_date TEXT NOT NULL, category TEXT NOT NULL, payee TEXT, description TEXT NOT NULL, reference TEXT, payment_method TEXT NOT NULL, account_id INTEGER, transaction_id INTEGER, amount REAL NOT NULL, created_at TEXT NOT NULL)",
    ),
    db.prepare(
      "CREATE TABLE IF NOT EXISTS document_sequences (kind TEXT PRIMARY KEY, prefix TEXT NOT NULL, last_number INTEGER NOT NULL DEFAULT 0, width INTEGER NOT NULL DEFAULT 4)",
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
  if (!invoiceCols.includes("stock_effects_json"))
    await addColumn(
      "invoices",
      "stock_effects_json",
      "TEXT NOT NULL DEFAULT '[]'",
    );
  const purchaseCols = (
    await db.prepare("PRAGMA table_info(purchases)").all()
  ).results.map((x) => String((x as Record<string, unknown>).name));
  for (const name of ["supplier_code", "item_code", "received_date"])
    if (!purchaseCols.includes(name))
      await addColumn("purchases", name, "TEXT");
  const userCols = (
    await db.prepare("PRAGMA table_info(users)").all()
  ).results.map((x) => String((x as Record<string, unknown>).name));
  for (const [name, type] of Object.entries({
    user_id: "TEXT",
    password_hash: "TEXT",
    password_salt: "TEXT",
    phone: "TEXT",
    designation: "TEXT",
    company: "TEXT",
    timezone: "TEXT",
    failed_login_count: "INTEGER NOT NULL DEFAULT 0",
    locked_until: "TEXT",
    last_login_at: "TEXT",
    created_at: "TEXT",
  }))
    if (!userCols.includes(name)) await addColumn("users", name, type);
  await db
    .prepare(
      "UPDATE users SET created_at = datetime('now') WHERE created_at IS NULL OR created_at = ''",
    )
    .run();
  await db
    .prepare(
      "CREATE UNIQUE INDEX IF NOT EXISTS users_user_id_unique ON users(user_id) WHERE user_id IS NOT NULL AND trim(user_id) != ''",
    )
    .run();
  const paymentCols = (
    await db.prepare("PRAGMA table_info(payments)").all()
  ).results.map((x) => String((x as Record<string, unknown>).name));
  for (const name of ["account_id", "transaction_id"])
    if (!paymentCols.includes(name))
      await addColumn("payments", name, "INTEGER");
  const returnCols = (
    await db.prepare("PRAGMA table_info(sales_returns)").all()
  ).results.map((x) => String((x as Record<string, unknown>).name));
  if (!returnCols.includes("items_json"))
    await addColumn(
      "sales_returns",
      "items_json",
      "TEXT NOT NULL DEFAULT '[]'",
    );
  if (!returnCols.includes("stock_effects_json"))
    await addColumn(
      "sales_returns",
      "stock_effects_json",
      "TEXT NOT NULL DEFAULT '[]'",
    );
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
export type AuthUser = {
  id: number;
  name: string;
  user_id: string;
  email?: string | null;
  role: string;
  enabled: number;
  phone?: string | null;
  designation?: string | null;
  company?: string | null;
  timezone?: string | null;
};
export async function getCurrentUser(request: Request) {
  return (await getSessionUser(request)) as AuthUser | null;
}
export async function authorize(request: Request) {
  return Boolean(await getCurrentUser(request));
}
const isViewer = (user: AuthUser) =>
  user.role.trim().toLowerCase() === "viewer";
const allowedRoles = new Set([
  "Admin",
  "Billing Staff",
  "Inventory Manager",
  "Viewer",
]);
const safeRole = (value: unknown) => {
  const role = String(value || "Viewer").trim();
  if (!allowedRoles.has(role)) throw new Error("Invalid user role");
  return role;
};
const canWrite = (user: AuthUser, resource: string) =>
  !isViewer(user) &&
  (resource !== "users" || user.role.trim().toLowerCase() === "admin");
function forbidden() {
  return Response.json(
    { error: "You do not have permission for this action" },
    { status: 403 },
  );
}
function mutationErrorStatus(message: string) {
  if (/unique constraint|already exists|already in use/i.test(message))
    return 409;
  if (
    /required|select |invalid|not found|not linked|quantity|available|requested|return amount|password|user id|at least one/i.test(
      message,
    )
  )
    return 400;
  return 500;
}
function clean(body: Record<string, unknown>, fields: string[]) {
  return Object.fromEntries(
    fields.filter((f) => body[f] !== undefined).map((f) => [f, body[f]]),
  );
}
function validateResourceQuantity(
  resource: string,
  body: Record<string, unknown>,
) {
  const field = resource === "products" ? "stock" : "quantity";
  if (
    (resource !== "products" && resource !== "purchases") ||
    body[field] === undefined
  )
    return;
  const quantity = Number(body[field]);
  if (
    !Number.isInteger(quantity) ||
    (resource === "products" ? quantity < 0 : quantity <= 0)
  )
    throw new Error(
      `${resource === "products" ? "Stock" : "Purchase quantity"} must be a whole number ${resource === "products" ? "of zero or more" : "above zero"}`,
    );
}
type ProductStockRow = {
  id: number;
  name: string;
  sku: string;
  stock: number;
};
type StockEffect = {
  product_id: number;
  sku: string;
  name: string;
  quantity: number;
};
const stockKey = (value: unknown) =>
  String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
const parseStockEffects = (value: unknown): StockEffect[] => {
  try {
    const parsed = JSON.parse(String(value || "[]"));
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map((effect) => ({
        product_id: Number(effect.product_id || 0),
        sku: String(effect.sku || ""),
        name: String(effect.name || ""),
        quantity: Number(effect.quantity || 0),
      }))
      .filter((effect) => effect.product_id && effect.quantity > 0);
  } catch {
    return [];
  }
};
const parseInvoiceItems = (value: unknown) => {
  try {
    const parsed = JSON.parse(String(value || "[]"));
    return Array.isArray(parsed) ? (parsed as Record<string, unknown>[]) : [];
  } catch {
    throw new Error("Invoice items are damaged or invalid");
  }
};
function applyInvoiceStockPlan(
  invoice: Record<string, unknown>,
  products: ProductStockRow[],
  stockById: Map<number, number>,
  previousEffects: StockEffect[] = [],
) {
  for (const effect of previousEffects)
    stockById.set(
      effect.product_id,
      Number(stockById.get(effect.product_id) || 0) + effect.quantity,
    );
  if (String(invoice.kind || "invoice").toLowerCase() !== "invoice")
    return [] as StockEffect[];
  const items = parseInvoiceItems(invoice.items_json);
  if (!items.length) throw new Error("Add at least one invoice item");
  const byCode = new Map(
      products.map((product) => [stockKey(product.sku), product]),
    ),
    byName = new Map(
      products.map((product) => [stockKey(product.name), product]),
    ),
    quantities = new Map<
      number,
      { product: ProductStockRow; quantity: number }
    >(),
    allowUnmatched = invoice.allow_unmatched_stock === true;
  for (const item of items) {
    const quantity = Number(item.quantity ?? item.qty ?? 0),
      code = item.code ?? item.sku ?? item.item_code,
      name = item.item ?? item.name ?? item.description,
      product =
        (stockKey(code) ? byCode.get(stockKey(code)) : undefined) ||
        (!stockKey(code) && stockKey(name)
          ? byName.get(stockKey(name))
          : undefined);
    if (!Number.isFinite(quantity) || quantity <= 0)
      throw new Error(
        `Enter a valid quantity for ${String(name || code || "the item")}`,
      );
    if (!Number.isInteger(quantity))
      throw new Error(
        `Quantity for ${String(name || code || "the item")} must be a whole number`,
      );
    if (!product) {
      if (allowUnmatched) continue;
      throw new Error(
        `${String(name || code || "Invoice item")} is not linked to Stock. Select the item from the Stock list.`,
      );
    }
    const existing = quantities.get(product.id);
    quantities.set(product.id, {
      product,
      quantity: Number(existing?.quantity || 0) + quantity,
    });
  }
  const effects: StockEffect[] = [];
  for (const { product, quantity } of quantities.values()) {
    const available = Number(stockById.get(product.id) || 0);
    if (quantity > available + 0.000001)
      throw new Error(
        `${product.name} has only ${available} available; ${quantity} was requested`,
      );
    stockById.set(product.id, available - quantity);
    effects.push({
      product_id: product.id,
      sku: product.sku,
      name: product.name,
      quantity,
    });
  }
  return effects;
}
async function loadProductStock() {
  return (
    await getEnv()
      .DB.prepare("SELECT id, name, sku, stock FROM products ORDER BY id")
      .all<ProductStockRow>()
  ).results;
}
type DocumentKind = "invoice" | "return";
const documentParts = (value: unknown) => {
  const match = String(value || "")
    .trim()
    .match(/^(.*?)(\d+)$/);
  return match
    ? {
        prefix: match[1],
        number: Number(match[2]),
        width: match[2].length,
      }
    : null;
};
export async function nextDocumentNumber(kind: DocumentKind) {
  const db = getEnv().DB;
  let sequence = await db
    .prepare(
      "SELECT prefix, last_number, width FROM document_sequences WHERE kind = ? LIMIT 1",
    )
    .bind(kind)
    .first<{ prefix: string; last_number: number; width: number }>();
  if (!sequence) {
    const rows = (
        kind === "invoice"
          ? await db
              .prepare(
                "SELECT invoice_no AS document_no FROM invoices WHERE kind = 'invoice' ORDER BY id DESC LIMIT 10000",
              )
              .all<{ document_no: string }>()
          : await db
              .prepare(
                "SELECT credit_note_no AS document_no FROM sales_returns ORDER BY id DESC LIMIT 10000",
              )
              .all<{ document_no: string }>()
      ).results,
      existing = rows
        .map((row) => documentParts(row.document_no))
        .find(Boolean),
      settings =
        kind === "invoice"
          ? await db
              .prepare("SELECT prefix FROM settings ORDER BY id LIMIT 1")
              .first<{ prefix: string }>()
          : null;
    sequence = {
      prefix:
        existing?.prefix ||
        settings?.prefix ||
        (kind === "invoice" ? "INV-" : "CN-"),
      last_number: existing?.number || 0,
      width: existing?.width || 4,
    };
    await db
      .prepare(
        "INSERT OR IGNORE INTO document_sequences(kind,prefix,last_number,width) VALUES(?,?,?,?)",
      )
      .bind(kind, sequence.prefix, sequence.last_number, sequence.width)
      .run();
  }
  return `${sequence.prefix}${String(Number(sequence.last_number || 0) + 1).padStart(Number(sequence.width || 4), "0")}`;
}
export async function advanceDocumentSequence(
  kind: DocumentKind,
  documentNumber: unknown,
) {
  const parts = documentParts(documentNumber);
  if (!parts) return;
  const db = getEnv().DB,
    current = await db
      .prepare(
        "SELECT prefix, last_number, width FROM document_sequences WHERE kind = ? LIMIT 1",
      )
      .bind(kind)
      .first<{ prefix: string; last_number: number; width: number }>();
  if (!current || current.prefix !== parts.prefix) {
    await db
      .prepare(
        "INSERT INTO document_sequences(kind,prefix,last_number,width) VALUES(?,?,?,?) ON CONFLICT(kind) DO UPDATE SET prefix=excluded.prefix,last_number=excluded.last_number,width=excluded.width",
      )
      .bind(kind, parts.prefix, parts.number, parts.width)
      .run();
    return;
  }
  await db
    .prepare(
      "UPDATE document_sequences SET last_number = MAX(last_number, ?), width = MAX(width, ?) WHERE kind = ?",
    )
    .bind(parts.number, parts.width, kind)
    .run();
}
async function validateSalesReturn(
  body: Record<string, unknown>,
  currentId = 0,
  pendingByInvoice = new Map<number, number>(),
) {
  const db = getEnv().DB,
    invoiceId = Number(body.invoice_id || 0),
    amount = Number(body.amount || 0);
  if (!invoiceId) throw new Error("Select the sales invoice for this return");
  if (!Number.isFinite(amount) || amount <= 0)
    throw new Error("Return amount must be greater than zero");
  const invoice = await db
    .prepare(
      "SELECT id, customer_id, customer_name, amount, kind FROM invoices WHERE id = ? LIMIT 1",
    )
    .bind(invoiceId)
    .first<Record<string, unknown>>();
  if (!invoice || String(invoice.kind).toLowerCase() !== "invoice")
    throw new Error("The selected sales invoice was not found");
  const existing = await db
      .prepare(
        "SELECT COALESCE(SUM(amount), 0) AS amount FROM sales_returns WHERE invoice_id = ? AND id != ?",
      )
      .bind(invoiceId, currentId)
      .first<{ amount: number }>(),
    returned =
      Number(existing?.amount || 0) +
      Number(pendingByInvoice.get(invoiceId) || 0);
  if (returned + amount > Number(invoice.amount || 0) + 0.000001)
    throw new Error(
      `Return amount exceeds the remaining value of invoice ${String(invoice.id)}`,
    );
  body.customer_id = Number(invoice.customer_id);
  pendingByInvoice.set(
    invoiceId,
    Number(pendingByInvoice.get(invoiceId) || 0) + amount,
  );
}
const returnItemKey = (item: Record<string, unknown>) => {
  const code = stockKey(item.code ?? item.sku ?? item.item_code),
    name = stockKey(item.item ?? item.name ?? item.description);
  return code ? `code:${code}` : `name:${name}`;
};
async function validateInvoiceReturnedQuantities(
  invoice: Record<string, unknown>,
  invoiceId: number,
) {
  const returnedRows = (
    await getEnv()
      .DB.prepare(
        "SELECT items_json FROM sales_returns WHERE invoice_id = ? AND items_json != '[]'",
      )
      .bind(invoiceId)
      .all<Record<string, unknown>>()
  ).results;
  if (!returnedRows.length) return;
  if (String(invoice.kind || "invoice").toLowerCase() !== "invoice")
    throw new Error(
      "Delete the linked sales returns before converting this invoice",
    );
  const sold = new Map<string, number>(),
    returned = new Map<string, number>();
  for (const item of parseInvoiceItems(invoice.items_json)) {
    const key = returnItemKey(item);
    sold.set(
      key,
      Number(sold.get(key) || 0) + Number(item.quantity ?? item.qty ?? 0),
    );
  }
  for (const row of returnedRows)
    for (const item of parseInvoiceItems(row.items_json)) {
      const key = returnItemKey(item);
      returned.set(
        key,
        Number(returned.get(key) || 0) + Number(item.quantity ?? item.qty ?? 0),
      );
    }
  for (const [key, quantity] of returned)
    if (quantity > Number(sold.get(key) || 0))
      throw new Error(
        "Invoice quantity cannot be reduced below the quantity already returned",
      );
}
async function prepareItemSalesReturn(
  body: Record<string, unknown>,
  currentId = 0,
) {
  const requestedItems = parseInvoiceItems(body.items_json);
  if (!requestedItems.length) {
    await validateSalesReturn(body, currentId);
    return null;
  }
  const db = getEnv().DB,
    invoiceId = Number(body.invoice_id || 0),
    invoice = await db
      .prepare(
        "SELECT id, customer_id, customer_name, amount, kind, items_json FROM invoices WHERE id = ? LIMIT 1",
      )
      .bind(invoiceId)
      .first<Record<string, unknown>>();
  if (!invoice || String(invoice.kind).toLowerCase() !== "invoice")
    throw new Error("The selected sales invoice was not found");
  const soldItems = parseInvoiceItems(invoice.items_json),
    soldByKey = new Map<
      string,
      {
        key: string;
        code: string;
        name: string;
        quantity: number;
        rate: number;
        gst: number;
      }
    >();
  for (const item of soldItems) {
    const key = returnItemKey(item),
      quantity = Number(item.quantity ?? item.qty ?? 0),
      existing = soldByKey.get(key);
    if (!key.replace(/^\w+:/, "") || quantity <= 0) continue;
    soldByKey.set(key, {
      key,
      code: String(
        item.code ?? item.sku ?? item.item_code ?? existing?.code ?? "",
      ),
      name: String(
        item.item ?? item.name ?? item.description ?? existing?.name ?? "",
      ),
      quantity: Number(existing?.quantity || 0) + quantity,
      rate: Number(item.rate ?? item.unit_price ?? existing?.rate ?? 0),
      gst: Number(item.gst ?? item.gst_rate ?? existing?.gst ?? 0),
    });
  }
  const previousReturns = (
      await db
        .prepare(
          "SELECT id, items_json FROM sales_returns WHERE invoice_id = ? AND id != ?",
        )
        .bind(invoiceId, currentId)
        .all<Record<string, unknown>>()
    ).results,
    previouslyReturned = new Map<string, number>();
  for (const returned of previousReturns)
    for (const item of parseInvoiceItems(returned.items_json)) {
      const key = returnItemKey(item);
      previouslyReturned.set(
        key,
        Number(previouslyReturned.get(key) || 0) +
          Number(item.quantity ?? item.qty ?? 0),
      );
    }
  const products = await loadProductStock(),
    initialStock = new Map(
      products.map((product) => [product.id, Number(product.stock || 0)]),
    ),
    stockById = new Map(initialStock),
    current = currentId
      ? await db
          .prepare(
            "SELECT stock_effects_json FROM sales_returns WHERE id = ? LIMIT 1",
          )
          .bind(currentId)
          .first<Record<string, unknown>>()
      : null;
  for (const effect of parseStockEffects(current?.stock_effects_json)) {
    const available = Number(stockById.get(effect.product_id) || 0);
    if (available < effect.quantity)
      throw new Error(
        `${effect.name || effect.sku} stock was already sold and this return cannot be changed`,
      );
    stockById.set(effect.product_id, available - effect.quantity);
  }
  const byCode = new Map(
      products.map((product) => [stockKey(product.sku), product]),
    ),
    byName = new Map(
      products.map((product) => [stockKey(product.name), product]),
    ),
    requestedByKey = new Map<string, number>();
  for (const item of requestedItems) {
    const key = returnItemKey(item),
      sold = soldByKey.get(key),
      quantity = Number(item.quantity ?? item.qty ?? 0);
    if (!sold) throw new Error("A selected return item is not in this invoice");
    if (!Number.isInteger(quantity) || quantity <= 0)
      throw new Error(
        `Return quantity for ${sold.name} must be a whole number above zero`,
      );
    requestedByKey.set(key, Number(requestedByKey.get(key) || 0) + quantity);
  }
  const effectsByProduct = new Map<number, StockEffect>(),
    normalizedItems: Record<string, unknown>[] = [];
  let amount = 0;
  for (const [key, quantity] of requestedByKey) {
    const sold = soldByKey.get(key)!,
      remaining = sold.quantity - Number(previouslyReturned.get(key) || 0);
    if (quantity > remaining)
      throw new Error(
        `${sold.name} has only ${remaining} quantity available to return`,
      );
    const product =
      (sold.code ? byCode.get(stockKey(sold.code)) : undefined) ||
      byName.get(stockKey(sold.name));
    if (!product)
      throw new Error(`${sold.name} is not linked to Stock by product code`);
    const lineAmount = quantity * sold.rate * (1 + sold.gst / 100),
      existingEffect = effectsByProduct.get(product.id);
    amount += lineAmount;
    normalizedItems.push({
      code: sold.code || product.sku,
      name: sold.name || product.name,
      quantity,
      rate: sold.rate,
      gst: sold.gst,
      total: Math.round(lineAmount * 100) / 100,
    });
    effectsByProduct.set(product.id, {
      product_id: product.id,
      sku: product.sku,
      name: product.name,
      quantity: Number(existingEffect?.quantity || 0) + quantity,
    });
  }
  if (!normalizedItems.length)
    throw new Error("Select at least one item and enter its return quantity");
  for (const effect of effectsByProduct.values())
    stockById.set(
      effect.product_id,
      Number(stockById.get(effect.product_id) || 0) + effect.quantity,
    );
  body.customer_id = Number(invoice.customer_id);
  body.amount = Math.round(amount * 100) / 100;
  body.items_json = JSON.stringify(normalizedItems);
  return {
    products,
    initialStock,
    stockById,
    effects: [...effectsByProduct.values()],
  };
}
async function reverseAndDeleteReturns(id?: number) {
  const db = getEnv().DB,
    rows = (
      id
        ? await db
            .prepare(
              "SELECT id, stock_effects_json FROM sales_returns WHERE id = ?",
            )
            .bind(id)
            .all<Record<string, unknown>>()
        : await db
            .prepare("SELECT id, stock_effects_json FROM sales_returns")
            .all<Record<string, unknown>>()
    ).results,
    removeByProduct = new Map<number, { quantity: number; name: string }>();
  for (const row of rows)
    for (const effect of parseStockEffects(row.stock_effects_json)) {
      const existing = removeByProduct.get(effect.product_id);
      removeByProduct.set(effect.product_id, {
        quantity: Number(existing?.quantity || 0) + effect.quantity,
        name: effect.name || effect.sku,
      });
    }
  for (const [productId, effect] of removeByProduct) {
    const product = await db
      .prepare("SELECT stock FROM products WHERE id = ? LIMIT 1")
      .bind(productId)
      .first<{ stock: number }>();
    if (Number(product?.stock || 0) < effect.quantity)
      throw new Error(
        `${effect.name} returned stock has already been sold; this credit note cannot be deleted`,
      );
  }
  const stockStatements = [...removeByProduct.entries()].map(
      ([productId, effect]) =>
        db
          .prepare("UPDATE products SET stock = stock - ? WHERE id = ?")
          .bind(effect.quantity, productId),
    ),
    remove = id
      ? db.prepare("DELETE FROM sales_returns WHERE id = ?").bind(id)
      : db.prepare("DELETE FROM sales_returns");
  await db.batch([...stockStatements, remove]);
}
async function connectPaymentToInvoice(body: Record<string, unknown>) {
  const invoiceId = Number(body.invoice_id || 0);
  if (!invoiceId) return;
  const invoice = await getEnv()
    .DB.prepare(
      "SELECT id, customer_id, kind FROM invoices WHERE id = ? LIMIT 1",
    )
    .bind(invoiceId)
    .first<Record<string, unknown>>();
  if (!invoice || String(invoice.kind).toLowerCase() !== "invoice")
    throw new Error("The selected sales invoice was not found");
  body.customer_id = Number(invoice.customer_id);
}
async function restoreAndDeleteInvoices(
  options: {
    id?: number;
    kind?: string;
  } = {},
) {
  const db = getEnv().DB,
    clause = options.id ? "WHERE id = ?" : options.kind ? "WHERE kind = ?" : "",
    binding = options.id || options.kind,
    prepared = db.prepare(
      `SELECT id, stock_effects_json FROM invoices ${clause}`,
    ),
    rows = (
      binding === undefined
        ? await prepared.all()
        : await prepared.bind(binding).all()
    ).results as Record<string, unknown>[],
    invoiceIds = rows.map((row) => Number(row.id)),
    linkedReturns = invoiceIds.length
      ? (
          await db
            .prepare(
              `SELECT stock_effects_json FROM sales_returns WHERE invoice_id IN (${invoiceIds.map(() => "?").join(",")})`,
            )
            .bind(...invoiceIds)
            .all<Record<string, unknown>>()
        ).results
      : [],
    restoreByProduct = new Map<number, number>();
  for (const row of rows)
    for (const effect of parseStockEffects(row.stock_effects_json))
      restoreByProduct.set(
        effect.product_id,
        Number(restoreByProduct.get(effect.product_id) || 0) + effect.quantity,
      );
  for (const row of linkedReturns)
    for (const effect of parseStockEffects(row.stock_effects_json))
      restoreByProduct.set(
        effect.product_id,
        Number(restoreByProduct.get(effect.product_id) || 0) - effect.quantity,
      );
  for (const [productId, quantity] of restoreByProduct) {
    const product = await db
      .prepare("SELECT stock, name FROM products WHERE id = ? LIMIT 1")
      .bind(productId)
      .first<{ stock: number; name: string }>();
    if (Number(product?.stock || 0) + quantity < 0)
      throw new Error(
        `${product?.name || "Returned item"} stock would become negative; linked returned stock was already sold`,
      );
  }
  const restoreStatements = [...restoreByProduct.entries()].map(
      ([productId, quantity]) =>
        db
          .prepare("UPDATE products SET stock = stock + ? WHERE id = ?")
          .bind(quantity, productId),
    ),
    deleteReturns = invoiceIds.length
      ? db
          .prepare(
            `DELETE FROM sales_returns WHERE invoice_id IN (${invoiceIds.map(() => "?").join(",")})`,
          )
          .bind(...invoiceIds)
      : null,
    deleteStatement = db.prepare(`DELETE FROM invoices ${clause}`);
  await db.batch([
    ...restoreStatements,
    ...(deleteReturns ? [deleteReturns] : []),
    binding === undefined ? deleteStatement : deleteStatement.bind(binding),
  ]);
}
async function linkLedgerMovement(
  resource: "payments" | "expenses",
  row: Record<string, unknown>,
) {
  const db = getEnv().DB,
    accountId = Number(row.account_id);
  if (!accountId) throw new Error("A cash or bank account is required");
  const account = await db
    .prepare("SELECT id, name, type FROM accounts WHERE id = ? LIMIT 1")
    .bind(accountId)
    .first<Record<string, unknown>>();
  if (!account)
    throw new Error("The selected cash or bank account was not found");
  let particulars = String(row.description || ""),
    reference = String(row.reference || ""),
    direction = "out",
    transactionDate = String(row.expense_date || "");
  if (resource === "payments") {
    const customer = await db
        .prepare("SELECT name FROM customers WHERE id = ? LIMIT 1")
        .bind(Number(row.customer_id))
        .first<{ name: string }>(),
      invoice = row.invoice_id
        ? await db
            .prepare("SELECT invoice_no FROM invoices WHERE id = ? LIMIT 1")
            .bind(Number(row.invoice_id))
            .first<{ invoice_no: string }>()
        : null;
    direction = "in";
    transactionDate = String(row.collected_at || "");
    particulars = `Payment from ${customer?.name || "customer"}`;
    reference = reference || invoice?.invoice_no || `PAY-${row.id}`;
  } else {
    particulars = `Expense: ${String(row.description || row.category || "Operating expense")}`;
    reference = reference || `EXP-${row.id}`;
  }
  const transaction = await db
    .prepare(
      "INSERT INTO account_transactions(account_id,direction,amount,particulars,reference,transaction_date) VALUES(?,?,?,?,?,?) RETURNING id",
    )
    .bind(
      accountId,
      direction,
      Number(row.amount || 0),
      particulars,
      reference,
      transactionDate,
    )
    .first<{ id: number }>();
  if (!transaction?.id)
    throw new Error("Unable to create the linked account entry");
  await db
    .prepare(`UPDATE ${resource} SET transaction_id = ? WHERE id = ?`)
    .bind(transaction.id, Number(row.id))
    .run();
  return { ...row, transaction_id: transaction.id };
}
export async function GET(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user)
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 401 },
      );
    const c = cfg(request);
    if (isViewer(user) && c.key !== "products") return forbidden();
    if (c.key === "users" && user.role.trim().toLowerCase() !== "admin")
      return forbidden();
    if (isViewer(user)) {
      const rows = await getEnv()
        .DB.prepare(
          "SELECT id, name, stock, price FROM products ORDER BY name LIMIT 10000",
        )
        .all();
      return Response.json({ rows: rows.results });
    }
    const rows = await getEnv()
      .DB.prepare(
        c.key === "users"
          ? `SELECT id,name,user_id,role,enabled,phone,designation,company,timezone,last_login_at,created_at
             FROM users ORDER BY id DESC LIMIT 10000`
          : `SELECT * FROM ${c.table} ORDER BY id DESC LIMIT 10000`,
      )
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
    const user = await getCurrentUser(request);
    if (!user)
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 401 },
      );
    const c = cfg(request);
    if (!canWrite(user, c.key)) return forbidden();
    const payload = (await request.json()) as
      | Record<string, unknown>
      | Record<string, unknown>[];
    const stampProduct = (row: Record<string, unknown>) => {
      if (c.key === "products" && !row.created_at)
        row.created_at = new Date().toISOString();
    };
    if (Array.isArray(payload)) {
      const db = getEnv().DB;
      if (c.key === "users")
        return Response.json(
          { error: "Users must be added individually" },
          { status: 400 },
        );
      payload.forEach(stampProduct);
      payload.forEach((body) => validateResourceQuantity(c.key, body));
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
      if (c.key === "returns") {
        const pendingByInvoice = new Map<number, number>();
        for (const body of payload)
          await validateSalesReturn(body, 0, pendingByInvoice);
      }
      if (c.key === "payments")
        for (const body of payload) await connectPaymentToInvoice(body);
      if (c.key === "invoices") {
        const products = await loadProductStock(),
          initialStock = new Map(
            products.map((product) => [product.id, Number(product.stock || 0)]),
          ),
          stockById = new Map(initialStock),
          existingNumbers = new Set(
            (
              await db.prepare("SELECT invoice_no FROM invoices").all<{
                invoice_no: string;
              }>()
            ).results.map((row) => String(row.invoice_no).toLowerCase()),
          ),
          invoiceStatements = [],
          savedNumbers: string[] = [];
        let count = 0;
        for (const body of payload) {
          for (const field of c.required)
            if (body[field] === undefined || body[field] === "")
              throw new Error(`${field} is required`);
          const invoiceNumber = String(body.invoice_no || "").toLowerCase();
          if (existingNumbers.has(invoiceNumber)) continue;
          const effects = applyInvoiceStockPlan(body, products, stockById),
            values: Record<string, unknown> = clean(body, c.fields);
          values.stock_effects_json = JSON.stringify(effects);
          const keys = Object.keys(values);
          invoiceStatements.push(
            db
              .prepare(
                `INSERT INTO invoices (${keys.join(",")}) VALUES (${keys.map(() => "?").join(",")})`,
              )
              .bind(...keys.map((key) => values[key])),
          );
          existingNumbers.add(invoiceNumber);
          if (
            body.allow_unmatched_stock !== true &&
            String(body.kind || "invoice").toLowerCase() === "invoice"
          )
            savedNumbers.push(String(body.invoice_no));
          count += 1;
        }
        const stockStatements = products
          .filter(
            (product) =>
              Number(initialStock.get(product.id)) !==
              Number(stockById.get(product.id)),
          )
          .map((product) =>
            db
              .prepare("UPDATE products SET stock = ? WHERE id = ?")
              .bind(Number(stockById.get(product.id)), product.id),
          );
        if (invoiceStatements.length || stockStatements.length)
          await db.batch([...invoiceStatements, ...stockStatements]);
        for (const number of savedNumbers)
          await advanceDocumentSequence("invoice", number);
        return Response.json({ count }, { status: 201 });
      }
      if (c.key === "payments" || c.key === "expenses") {
        let count = 0;
        for (const body of payload) {
          for (const field of c.required)
            if (body[field] === undefined || body[field] === "")
              throw new Error(`${field} is required`);
          const values = clean(body, c.fields),
            keys = Object.keys(values),
            created = await db
              .prepare(
                `INSERT INTO ${c.table} (${keys.join(",")}) VALUES (${keys.map(() => "?").join(",")}) RETURNING *`,
              )
              .bind(...keys.map((key) => values[key]))
              .first<Record<string, unknown>>();
          if (!created) throw new Error("Unable to import the account entry");
          try {
            await linkLedgerMovement(c.key, created);
            count += 1;
          } catch (error) {
            await db
              .prepare(`DELETE FROM ${c.table} WHERE id = ?`)
              .bind(Number(created.id))
              .run();
            throw error;
          }
        }
        return Response.json({ count }, { status: 201 });
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
    validateResourceQuantity(c.key, body);
    if (c.key === "users") {
      if (!String(body.name || "").trim())
        return Response.json({ error: "Name is required" }, { status: 400 });
      const userId = validateUserId(body.user_id),
        duplicate = await getEnv()
          .DB.prepare(
            "SELECT id FROM users WHERE lower(user_id) = lower(?) LIMIT 1",
          )
          .bind(userId)
          .first<{ id: number }>();
      if (duplicate)
        return Response.json(
          { error: "That User ID is already in use" },
          { status: 409 },
        );
      const password = await createPassword(body.password);
      body.user_id = userId;
      body.email = `${userId}@billflow.local`;
      body.role = safeRole(body.role);
      body.enabled = body.enabled === undefined ? 1 : Number(body.enabled);
      const values = {
          ...clean(body, c.fields),
          password_hash: password.hash,
          password_salt: password.salt,
          failed_login_count: 0,
          created_at: new Date().toISOString(),
        },
        keys = Object.keys(values),
        created = await getEnv()
          .DB.prepare(
            `INSERT INTO users (${keys.join(",")}) VALUES (${keys.map(() => "?").join(",")}) RETURNING id`,
          )
          .bind(...keys.map((key) => values[key as keyof typeof values]))
          .first<{ id: number }>(),
        row = created?.id
          ? await getEnv()
              .DB.prepare(
                `SELECT id,name,user_id,role,enabled,phone,designation,company,timezone,last_login_at,created_at
                 FROM users WHERE id = ?`,
              )
              .bind(created.id)
              .first()
          : null;
      return Response.json({ row }, { status: 201 });
    }
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
    if (c.key === "payments") await connectPaymentToInvoice(body);
    const returnPlan =
      c.key === "returns" ? await prepareItemSalesReturn(body) : null;
    for (const f of c.required)
      if (body[f] === undefined || body[f] === "")
        return Response.json({ error: `${f} is required` }, { status: 400 });
    if (c.key === "returns" && returnPlan) {
      const db = getEnv().DB,
        values: Record<string, unknown> = clean(body, c.fields);
      values.stock_effects_json = JSON.stringify(returnPlan.effects);
      const keys = Object.keys(values),
        insert = db
          .prepare(
            `INSERT INTO sales_returns (${keys.join(",")}) VALUES (${keys.map(() => "?").join(",")})`,
          )
          .bind(...keys.map((key) => values[key])),
        stockStatements = returnPlan.products
          .filter(
            (product) =>
              Number(returnPlan.initialStock.get(product.id)) !==
              Number(returnPlan.stockById.get(product.id)),
          )
          .map((product) =>
            db
              .prepare("UPDATE products SET stock = ? WHERE id = ?")
              .bind(Number(returnPlan.stockById.get(product.id)), product.id),
          );
      await db.batch([insert, ...stockStatements]);
      await advanceDocumentSequence("return", body.credit_note_no);
      const result = await db
        .prepare("SELECT * FROM sales_returns WHERE credit_note_no = ? LIMIT 1")
        .bind(String(body.credit_note_no))
        .first<Record<string, unknown>>();
      return Response.json({ row: result }, { status: 201 });
    }
    if (c.key === "invoices") {
      const db = getEnv().DB,
        products = await loadProductStock(),
        initialStock = new Map(
          products.map((product) => [product.id, Number(product.stock || 0)]),
        ),
        stockById = new Map(initialStock),
        effects = applyInvoiceStockPlan(body, products, stockById),
        values: Record<string, unknown> = clean(body, c.fields);
      values.stock_effects_json = JSON.stringify(effects);
      const keys = Object.keys(values),
        insert = db
          .prepare(
            `INSERT INTO invoices (${keys.join(",")}) VALUES (${keys.map(() => "?").join(",")})`,
          )
          .bind(...keys.map((key) => values[key])),
        stockStatements = products
          .filter(
            (product) =>
              Number(initialStock.get(product.id)) !==
              Number(stockById.get(product.id)),
          )
          .map((product) =>
            db
              .prepare("UPDATE products SET stock = ? WHERE id = ?")
              .bind(Number(stockById.get(product.id)), product.id),
          );
      await db.batch([insert, ...stockStatements]);
      if (String(body.kind || "invoice").toLowerCase() === "invoice")
        await advanceDocumentSequence("invoice", body.invoice_no);
      const result = await db
        .prepare("SELECT * FROM invoices WHERE invoice_no = ? LIMIT 1")
        .bind(String(body.invoice_no))
        .first<Record<string, unknown>>();
      return Response.json({ row: result }, { status: 201 });
    }
    const values = clean(body, c.fields),
      keys = Object.keys(values);
    let result = await getEnv()
      .DB.prepare(
        `INSERT INTO ${c.table} (${keys.join(",")}) VALUES (${keys.map(() => "?").join(",")}) RETURNING *`,
      )
      .bind(...keys.map((k) => values[k]))
      .first<Record<string, unknown>>();
    if (result && (c.key === "payments" || c.key === "expenses"))
      result = await linkLedgerMovement(c.key, result);
    if (result && c.key === "returns")
      await advanceDocumentSequence("return", result.credit_note_no);
    return Response.json({ row: result }, { status: 201 });
  } catch (e) {
    const message = e instanceof Error ? e.message : "Unable to save data";
    return Response.json(
      {
        error: /unique constraint/i.test(message)
          ? "That User ID or document number already exists"
          : message,
      },
      { status: mutationErrorStatus(message) },
    );
  }
}
export async function PATCH(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user)
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 401 },
      );
    const c = cfg(request);
    if (!canWrite(user, c.key)) return forbidden();
    const payload = (await request.json()) as
      | Record<string, unknown>
      | Record<string, unknown>[];
    if (Array.isArray(payload)) {
      if (c.key === "users")
        return Response.json(
          { error: "Users must be updated individually" },
          { status: 400 },
        );
      const db = getEnv().DB;
      payload.forEach((body) => validateResourceQuantity(c.key, body));
      if (c.key === "returns") {
        if (payload.some((body) => parseInvoiceItems(body.items_json).length))
          throw new Error(
            "Item-level sales returns must be updated individually",
          );
        const currentReturns = (
            await db.prepare("SELECT * FROM sales_returns").all()
          ).results as Record<string, unknown>[],
          pendingByInvoice = new Map<number, number>();
        for (const body of payload) {
          const id = Number(body.id),
            current = currentReturns.find((row) => Number(row.id) === id);
          if (!id || !current) throw new Error("Sales return was not found");
          const merged = { ...current, ...body };
          await validateSalesReturn(merged, id, pendingByInvoice);
          body.invoice_id = merged.invoice_id;
          body.customer_id = merged.customer_id;
          body.amount = merged.amount;
        }
      }
      if (c.key === "invoices") {
        const products = await loadProductStock(),
          initialStock = new Map(
            products.map((product) => [product.id, Number(product.stock || 0)]),
          ),
          stockById = new Map(initialStock),
          currentInvoices = (await db.prepare("SELECT * FROM invoices").all())
            .results as Record<string, unknown>[],
          invoiceStatements = [],
          changedNumbers: unknown[] = [];
        for (const body of payload) {
          const id = Number(body.id),
            current = currentInvoices.find((row) => Number(row.id) === id);
          if (!id || !current) throw new Error("Invoice was not found");
          const merged = { ...current, ...body };
          await validateInvoiceReturnedQuantities(merged, id);
          const effects = applyInvoiceStockPlan(
              merged,
              products,
              stockById,
              parseStockEffects(current.stock_effects_json),
            ),
            values: Record<string, unknown> = clean(body, c.fields);
          values.stock_effects_json = JSON.stringify(effects);
          const keys = Object.keys(values);
          invoiceStatements.push(
            db
              .prepare(
                `UPDATE invoices SET ${keys.map((key) => `${key}=?`).join(",")} WHERE id=?`,
              )
              .bind(...keys.map((key) => values[key]), id),
          );
          if (
            body.invoice_no !== undefined &&
            String(body.invoice_no) !== String(current.invoice_no) &&
            String(merged.kind || "invoice").toLowerCase() === "invoice"
          )
            changedNumbers.push(body.invoice_no);
        }
        const stockStatements = products
          .filter(
            (product) =>
              Number(initialStock.get(product.id)) !==
              Number(stockById.get(product.id)),
          )
          .map((product) =>
            db
              .prepare("UPDATE products SET stock = ? WHERE id = ?")
              .bind(Number(stockById.get(product.id)), product.id),
          );
        if (invoiceStatements.length || stockStatements.length)
          await db.batch([...invoiceStatements, ...stockStatements]);
        for (const number of changedNumbers)
          await advanceDocumentSequence("invoice", number);
        return Response.json({ count: invoiceStatements.length });
      }
      const statements = payload.map((body) => {
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
    validateResourceQuantity(c.key, body);
    let shouldAdvanceReturn = false;
    if (c.key === "returns") {
      const current = await getEnv()
        .DB.prepare("SELECT * FROM sales_returns WHERE id = ? LIMIT 1")
        .bind(id)
        .first<Record<string, unknown>>();
      if (!current)
        return Response.json(
          { error: "Sales return not found" },
          { status: 404 },
        );
      const merged = { ...current, ...body },
        returnPlan = await prepareItemSalesReturn(merged, id);
      shouldAdvanceReturn =
        body.credit_note_no !== undefined &&
        String(body.credit_note_no) !== String(current.credit_note_no);
      body.invoice_id = merged.invoice_id;
      body.customer_id = merged.customer_id;
      body.amount = merged.amount;
      body.items_json = merged.items_json;
      if (returnPlan) {
        const db = getEnv().DB,
          values: Record<string, unknown> = clean(body, c.fields);
        values.stock_effects_json = JSON.stringify(returnPlan.effects);
        const keys = Object.keys(values),
          update = db
            .prepare(
              `UPDATE sales_returns SET ${keys.map((key) => `${key}=?`).join(",")} WHERE id=?`,
            )
            .bind(...keys.map((key) => values[key]), id),
          stockStatements = returnPlan.products
            .filter(
              (product) =>
                Number(returnPlan.initialStock.get(product.id)) !==
                Number(returnPlan.stockById.get(product.id)),
            )
            .map((product) =>
              db
                .prepare("UPDATE products SET stock = ? WHERE id = ?")
                .bind(Number(returnPlan.stockById.get(product.id)), product.id),
            );
        await db.batch([update, ...stockStatements]);
        if (shouldAdvanceReturn)
          await advanceDocumentSequence("return", body.credit_note_no);
        const row = await db
          .prepare("SELECT * FROM sales_returns WHERE id = ? LIMIT 1")
          .bind(id)
          .first<Record<string, unknown>>();
        return Response.json({ row });
      }
    }
    if (c.key === "invoices") {
      const db = getEnv().DB,
        current = await db
          .prepare("SELECT * FROM invoices WHERE id = ? LIMIT 1")
          .bind(id)
          .first<Record<string, unknown>>();
      if (!current)
        return Response.json({ error: "Invoice not found" }, { status: 404 });
      const products = await loadProductStock(),
        initialStock = new Map(
          products.map((product) => [product.id, Number(product.stock || 0)]),
        ),
        stockById = new Map(initialStock),
        merged = { ...current, ...body };
      await validateInvoiceReturnedQuantities(merged, id);
      const effects = applyInvoiceStockPlan(
          merged,
          products,
          stockById,
          parseStockEffects(current.stock_effects_json),
        ),
        values: Record<string, unknown> = clean(body, c.fields);
      values.stock_effects_json = JSON.stringify(effects);
      const keys = Object.keys(values),
        update = db
          .prepare(
            `UPDATE invoices SET ${keys.map((key) => `${key}=?`).join(",")} WHERE id=?`,
          )
          .bind(...keys.map((key) => values[key]), id),
        stockStatements = products
          .filter(
            (product) =>
              Number(initialStock.get(product.id)) !==
              Number(stockById.get(product.id)),
          )
          .map((product) =>
            db
              .prepare("UPDATE products SET stock = ? WHERE id = ?")
              .bind(Number(stockById.get(product.id)), product.id),
          );
      await db.batch([update, ...stockStatements]);
      if (
        body.invoice_no !== undefined &&
        String(body.invoice_no) !== String(current.invoice_no) &&
        String(merged.kind || "invoice").toLowerCase() === "invoice"
      )
        await advanceDocumentSequence("invoice", body.invoice_no);
      const row = await db
        .prepare("SELECT * FROM invoices WHERE id = ? LIMIT 1")
        .bind(id)
        .first<Record<string, unknown>>();
      return Response.json({ row });
    }
    if (c.key === "users") {
      const target = await getEnv()
        .DB.prepare("SELECT id, user_id, role, enabled FROM users WHERE id = ?")
        .bind(id)
        .first<{
          id: number;
          user_id: string;
          role: string;
          enabled: number;
        }>();
      if (!target)
        return Response.json({ error: "User not found" }, { status: 404 });
      const isSelf = target.id === user.id;
      if (isSelf && Number(body.enabled ?? target.enabled) === 0)
        return Response.json(
          {
            error:
              "Change your profile from My Profile; you cannot disable your own account",
          },
          { status: 400 },
        );
      const removingAdmin =
        target.role.toLowerCase() === "admin" &&
        (String(body.role || target.role).toLowerCase() !== "admin" ||
          Number(body.enabled ?? target.enabled) === 0);
      if (removingAdmin) {
        const admins = await getEnv()
          .DB.prepare(
            "SELECT COUNT(*) AS count FROM users WHERE lower(role) = 'admin' AND enabled = 1",
          )
          .first<{ count: number }>();
        if ((admins?.count || 0) <= 1)
          return Response.json(
            { error: "At least one active Admin user is required" },
            { status: 400 },
          );
      }
    }
    const values: Record<string, unknown> = clean(body, c.fields);
    if (c.key === "users") {
      if (body.user_id !== undefined)
        values.user_id = validateUserId(body.user_id);
      if (values.user_id !== undefined) {
        const duplicate = await getEnv()
          .DB.prepare(
            "SELECT id FROM users WHERE lower(user_id) = lower(?) AND id != ? LIMIT 1",
          )
          .bind(String(values.user_id), id)
          .first<{ id: number }>();
        if (duplicate)
          return Response.json(
            { error: "That User ID is already in use" },
            { status: 409 },
          );
      }
      if (body.role !== undefined) values.role = safeRole(body.role);
      if (body.password) {
        if (id === user.id)
          return Response.json(
            { error: "Change your own password from My Profile" },
            { status: 400 },
          );
        const password = await createPassword(body.password);
        values.password_hash = password.hash;
        values.password_salt = password.salt;
        values.failed_login_count = 0;
        values.locked_until = null;
      }
    }
    const keys = Object.keys(values);
    if (!keys.length)
      return Response.json({ error: "No fields to update" }, { status: 400 });
    let row = await getEnv()
      .DB.prepare(
        `UPDATE ${c.table} SET ${keys.map((k) => `${k}=?`).join(",")} WHERE id=? RETURNING *`,
      )
      .bind(...keys.map((k) => values[k]), id)
      .first<Record<string, unknown>>();
    if (c.key === "users") {
      if (Number(values.enabled) === 0 || body.password)
        await getEnv()
          .DB.prepare("DELETE FROM auth_sessions WHERE user_id = ?")
          .bind(id)
          .run();
      row = await getEnv()
        .DB.prepare(
          `SELECT id,name,user_id,role,enabled,phone,designation,company,timezone,last_login_at,created_at
           FROM users WHERE id = ?`,
        )
        .bind(id)
        .first<Record<string, unknown>>();
    }
    if (
      row &&
      (c.key === "payments" || c.key === "expenses") &&
      Number(row.transaction_id)
    ) {
      const isPayment = c.key === "payments";
      await getEnv()
        .DB.prepare(
          "UPDATE account_transactions SET account_id=?, direction=?, amount=?, particulars=?, reference=?, transaction_date=? WHERE id=?",
        )
        .bind(
          Number(row.account_id),
          isPayment ? "in" : "out",
          Number(row.amount),
          isPayment
            ? `Customer payment · ${String(row.reference || "")}`
            : `Expense: ${String(row.description || row.category || "")}`,
          String(row.reference || (isPayment ? `PAY-${id}` : `EXP-${id}`)),
          String(isPayment ? row.collected_at : row.expense_date),
          Number(row.transaction_id),
        )
        .run();
    }
    if (row && c.key === "returns" && shouldAdvanceReturn)
      await advanceDocumentSequence("return", row.credit_note_no);
    return Response.json({ row });
  } catch (e) {
    const message = e instanceof Error ? e.message : "Unable to update data";
    return Response.json(
      {
        error: /unique constraint/i.test(message)
          ? "That User ID is already in use"
          : message,
      },
      { status: mutationErrorStatus(message) },
    );
  }
}
export async function DELETE(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user)
      return Response.json(
        { error: "Unauthorized or disabled" },
        { status: 401 },
      );
    const c = cfg(request);
    if (!canWrite(user, c.key)) return forbidden();
    const url = new URL(request.url),
      id = Number(url.searchParams.get("id"));
    if (url.searchParams.get("all") === "1") {
      if (c.key === "users")
        return Response.json(
          { error: "Users must be deleted individually" },
          { status: 400 },
        );
      if (c.key === "payments" || c.key === "expenses") {
        await getEnv()
          .DB.prepare(
            `DELETE FROM account_transactions WHERE id IN (SELECT transaction_id FROM ${c.table} WHERE transaction_id IS NOT NULL)`,
          )
          .run();
      }
      if (c.key === "invoices" && url.searchParams.get("kind")) {
        await restoreAndDeleteInvoices({
          kind: String(url.searchParams.get("kind")),
        });
        return Response.json({ ok: true });
      }
      if (c.key === "returns") {
        await reverseAndDeleteReturns();
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
      if (c.key === "invoices") {
        await restoreAndDeleteInvoices();
        return Response.json({ ok: true });
      }
      await getEnv().DB.prepare(`DELETE FROM ${c.table}`).run();
      return Response.json({ ok: true });
    }
    if (!id) return Response.json({ error: "id is required" }, { status: 400 });
    if (c.key === "users") {
      const target = await getEnv()
        .DB.prepare("SELECT id, role, enabled FROM users WHERE id = ? LIMIT 1")
        .bind(id)
        .first<{ id: number; role: string; enabled: number }>();
      if (!target)
        return Response.json({ error: "User not found" }, { status: 404 });
      if (target.id === user.id)
        return Response.json(
          { error: "You cannot delete your own signed-in account" },
          { status: 400 },
        );
      if (target.role.toLowerCase() === "admin" && target.enabled) {
        const admins = await getEnv()
          .DB.prepare(
            "SELECT COUNT(*) AS count FROM users WHERE lower(role) = 'admin' AND enabled = 1",
          )
          .first<{ count: number }>();
        if ((admins?.count || 0) <= 1)
          return Response.json(
            { error: "At least one Admin user is required" },
            { status: 400 },
          );
      }
    }
    if (c.key === "users")
      await getEnv()
        .DB.prepare("DELETE FROM auth_sessions WHERE user_id = ?")
        .bind(id)
        .run();
    if (c.key === "payments" || c.key === "expenses") {
      const linked = await getEnv()
        .DB.prepare(`SELECT transaction_id FROM ${c.table} WHERE id = ?`)
        .bind(id)
        .first<{ transaction_id: number | null }>();
      if (linked?.transaction_id)
        await getEnv()
          .DB.prepare("DELETE FROM account_transactions WHERE id = ?")
          .bind(linked.transaction_id)
          .run();
    }
    if (c.key === "invoices") {
      await restoreAndDeleteInvoices({ id });
      return Response.json({ ok: true });
    }
    if (c.key === "returns") {
      await reverseAndDeleteReturns(id);
      return Response.json({ ok: true });
    }
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
