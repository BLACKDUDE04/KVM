"use client";
/* eslint-disable react-hooks/set-state-in-effect, react-hooks/exhaustive-deps */

import { useEffect, useMemo, useState } from "react";

type View =
  | "Dashboard"
  | "Invoices"
  | "Sales Returns"
  | "Purchases"
  | "Quotations"
  | "Customers"
  | "Suppliers"
  | "Inventory"
  | "Payments"
  | "Cash Account"
  | "Bank Accounts"
  | "Data Errors"
  | "Reports"
  | "Users"
  | "Settings";

const money = (n: number) =>
  new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(n);
const nav: [View, string][] = [
  ["Dashboard", "▦"],
  ["Invoices", "▤"],
  ["Sales Returns", "↩"],
  ["Purchases", "⇩"],
  ["Quotations", "▣"],
  ["Customers", "♧"],
  ["Suppliers", "♧"],
  ["Inventory", "◇"],
  ["Payments", "₹"],
  ["Cash Account", "₹"],
  ["Bank Accounts", "⌁"],
  ["Data Errors", "⚠"],
  ["Reports", "▥"],
  ["Users", "♙"],
  ["Settings", "⚙"],
];
type DataRow = Record<string, unknown> & { id: number };
const request = async (
  resource: string,
  method = "GET",
  body?: Record<string, unknown>,
  id?: number,
) => {
  const url = `/api/data?resource=${resource}${id ? `&id=${id}` : ""}`;
  const payload = method === "PATCH" && id ? { ...body, id } : body;
  const response = await fetch(url, {
    method,
    headers: payload ? { "content-type": "application/json" } : undefined,
    body: payload ? JSON.stringify(payload) : undefined,
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "Request failed");
  return data;
};
const clearResource = async (resource: string, filter = "") => {
  const response = await fetch(
    `/api/data?resource=${resource}&all=1${filter}`,
    {
      method: "DELETE",
    },
  );
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "Unable to clear data");
};
const exportExcel = async (rows: Record<string, unknown>[], name: string) => {
  const XLSX = await import("xlsx"),
    sheet = XLSX.utils.json_to_sheet(rows),
    book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "Data");
  XLSX.writeFile(book, `${name}-${new Date().toISOString().slice(0, 10)}.xlsx`);
};
const headerKey = (value: string) =>
  value === "#"
    ? "#"
    : value
        .replace(/\u00a0/g, " ")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, " ")
        .trim();
const readExcel = async (file: File) => {
  if (!file.size) throw new Error("The selected Excel file is empty.");
  const buffer = await file.arrayBuffer(),
    text = new TextDecoder().decode(buffer),
    looksLikeHtml = /<table[\s>]/i.test(text.slice(0, 5000));
  if (looksLikeHtml) {
    const document = new DOMParser().parseFromString(text, "text/html"),
      table = document.querySelector("table");
    if (!table)
      throw new Error("No data table was found in this Excel report.");
    const headers = [...table.querySelectorAll("thead th")].map(
      (node) => node.textContent?.trim() || "",
    );
    if (!headers.length)
      throw new Error("Excel report headings could not be read.");
    return [...table.querySelectorAll("tbody tr")]
      .map((tr) => {
        const values = [...tr.querySelectorAll("td")].map(
          (node) => node.textContent?.trim() || "",
        );
        return Object.fromEntries(
          headers.map((header, index) => [header, values[index] || ""]),
        );
      })
      .filter((row) =>
        Object.values(row).some((value) => String(value).trim()),
      );
  }
  const XLSX = await import("xlsx"),
    book = XLSX.read(buffer),
    sheet = book.Sheets[book.SheetNames[0]],
    raw = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
    }),
    expected = new Set(
      [
        "customer name",
        "customer code",
        "supplier name",
        "supplier code",
        "shop name",
        "owner name",
        "bill no",
        "bill date",
        "item code",
        "item name",
        "contact no",
        "account",
        "narration",
        "date",
        "qty",
        "product value",
        "unit price",
        "gstin",
        "stock",
        "dr",
        "cr",
      ].map(headerKey),
    ),
    headerRow = raw.findIndex(
      (row) =>
        row.filter((value) => expected.has(headerKey(String(value)))).length >=
        2,
    );
  if (headerRow < 0)
    throw new Error(
      "Excel headings not found. Please keep the supplied column headings in one row.",
    );
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: "",
    range: headerRow,
  });
};
const cell = (row: Record<string, unknown>, ...names: string[]) => {
  const wanted = names.map(headerKey),
    key = Object.keys(row).find((k) => wanted.includes(headerKey(k)));
  return key ? row[key] : "";
};
const excelNumber = (value: unknown) => {
  if (typeof value === "number") return value;
  const raw = String(value).trim(),
    negative = /^\(.*\)$/.test(raw),
    parsed = Number(raw.replace(/[^0-9.-]/g, ""));
  return Number.isFinite(parsed) ? (negative ? -Math.abs(parsed) : parsed) : 0;
};
const excelDate = (value: unknown) => {
  if (typeof value === "number") {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    return date.toISOString().slice(0, 10);
  }
  const raw = String(value).trim(),
    parts = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (parts) {
    const year = parts[3].length === 2 ? `20${parts[3]}` : parts[3];
    return `${year}-${parts[2].padStart(2, "0")}-${parts[1].padStart(2, "0")}`;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime())
    ? new Date().toISOString().slice(0, 10)
    : parsed.toISOString().slice(0, 10);
};
const formatDate = (value: unknown) => {
  const raw = String(value || "").trim();
  if (!raw) return "—";
  const ymd = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (ymd)
    return `${ymd[3].padStart(2, "0")}/${ymd[2].padStart(2, "0")}/${ymd[1]}`;
  const dmy = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (dmy)
    return `${dmy[1].padStart(2, "0")}/${dmy[2].padStart(2, "0")}/${dmy[3]}`;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime())
    ? raw
    : `${String(parsed.getDate()).padStart(2, "0")}/${String(parsed.getMonth() + 1).padStart(2, "0")}/${parsed.getFullYear()}`;
};
const resourceCache = new Map<string, DataRow[]>();
function useResource(resource: string) {
  const [rows, setRows] = useState<DataRow[]>(
      resourceCache.get(resource) || [],
    ),
    [loading, setLoading] = useState(!resourceCache.has(resource));
  const load = async () => {
    setLoading(true);
    try {
      const d = await request(resource);
      const next = d.rows || [];
      resourceCache.set(resource, next);
      setRows(next);
    } finally {
      setLoading(false);
    }
  };
  useEffect(() => {
    load();
  }, [resource]);
  return { rows, loading, reload: load };
}
const Empty = ({ text }: { text: string }) => (
  <div className="empty-import">{text}</div>
);
const DataControls = ({
  total,
  limit,
  setLimit,
  clear,
}: {
  total: number;
  limit: number;
  setLimit: (value: number) => void;
  clear: () => void;
}) => (
  <div className="data-controls">
    <span>
      Showing {Math.min(total, limit)} of {total} records
    </span>
    <label>
      View{" "}
      <select value={limit} onChange={(e) => setLimit(Number(e.target.value))}>
        <option value={25}>25</option>
        <option value={50}>50</option>
        <option value={100}>100</option>
        <option value={10000}>All</option>
      </select>
    </label>
    <button className="danger" onClick={clear}>
      Clear all data
    </button>
  </div>
);
const includesQuery = (row: unknown, query: string) =>
  !query.trim() ||
  JSON.stringify(row).toLowerCase().includes(query.trim().toLowerCase());
const normalizedValue = (value: unknown) =>
  String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
const rowFingerprint = (...values: unknown[]) =>
  values.map(normalizedValue).join("|");
const findProduct = (products: DataRow[], code: unknown, name?: unknown) => {
  const codeKey = headerKey(String(code || "")),
    nameKey = headerKey(String(name || ""));
  return products.find(
    (product) =>
      (codeKey && headerKey(String(product.sku || "")) === codeKey) ||
      (!codeKey &&
        nameKey &&
        headerKey(String(product.name || "")) === nameKey),
  );
};
const sameFields = (
  current: Record<string, unknown>,
  next: Record<string, unknown>,
  fields: string[],
) =>
  fields.every(
    (field) => normalizedValue(current[field]) === normalizedValue(next[field]),
  );
const billCompare = (a: unknown, b: unknown, direction: "asc" | "desc") =>
  String(a || "").localeCompare(String(b || ""), "en", {
    numeric: true,
    sensitivity: "base",
  }) * (direction === "asc" ? 1 : -1);
const buildPurchaseRateMap = (rows: DataRow[]) => {
  const rates = new Map<string, { rate: number; date: string; id: number }>();
  for (const purchase of rows) {
    const keys = [purchase.item_code, purchase.item]
        .map((value) => headerKey(String(value || "")))
        .filter(Boolean),
      quantity = Number(purchase.quantity || 0),
      rate =
        Number(purchase.rate || 0) ||
        (quantity ? Number(purchase.total || 0) / quantity : 0),
      date = String(purchase.purchase_date || "");
    for (const key of keys) {
      const current = rates.get(key);
      if (
        rate > 0 &&
        (!current ||
          date > current.date ||
          (date === current.date && purchase.id > current.id))
      )
        rates.set(key, { rate, date, id: purchase.id });
    }
  }
  return rates;
};
const effectivePurchaseRate = (
  product: DataRow,
  rates: Map<string, { rate: number }>,
) =>
  rates.get(headerKey(String(product.sku || "")))?.rate ||
  rates.get(headerKey(String(product.name || "")))?.rate ||
  Number(product.purchase_rate || 0);
function ModuleFilters({
  query,
  setQuery,
  placeholder,
  children,
}: {
  query: string;
  setQuery: (value: string) => void;
  placeholder: string;
  children?: React.ReactNode;
}) {
  return (
    <div className="module-filters">
      <label className="module-search">
        <span>⌕</span>
        <input
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder={placeholder}
          aria-label={placeholder}
        />
        {query && (
          <button type="button" onClick={() => setQuery("")}>
            Clear
          </button>
        )}
      </label>
      {children}
    </div>
  );
}
function BillDetails({
  title,
  party,
  date,
  items,
  close,
}: {
  title: string;
  party: string;
  date: string;
  items: Record<string, unknown>[];
  close: () => void;
}) {
  const total = items.reduce(
    (sum, item) => sum + Number(item.total || item.product_value || 0),
    0,
  );
  return (
    <div className="detail-backdrop" role="presentation" onClick={close}>
      <section
        className="bill-detail"
        role="dialog"
        aria-modal="true"
        aria-label={`${title} product details`}
        onClick={(event) => event.stopPropagation()}
      >
        <div className="form-head">
          <div>
            <small>PRODUCT DETAILS</small>
            <h3>{title}</h3>
            <p>
              {party} · {formatDate(date)}
            </p>
          </div>
          <button type="button" onClick={close} aria-label="Close details">
            ×
          </button>
        </div>
        <div className="excel-preview">
          <table>
            <thead>
              <tr>
                <th>Code</th>
                <th>Product</th>
                <th>Qty</th>
                <th>Rate</th>
                <th>Value</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item, index) => (
                <tr key={`${String(item.code || item.item_code)}-${index}`}>
                  <td>{String(item.code || item.item_code || "—")}</td>
                  <td>
                    <b>
                      {String(
                        item.item || item.name || item.description || "Item",
                      )}
                    </b>
                  </td>
                  <td>{Number(item.quantity || item.qty || 0)}</td>
                  <td>{money(Number(item.rate || item.unit_price || 0))}</td>
                  <td>
                    {money(Number(item.total || item.product_value || 0))}
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr>
                <td colSpan={4}>Bill total</td>
                <td>
                  <b>{money(total)}</b>
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </section>
    </div>
  );
}
type ImportStatus = {
  label: string;
  done: number;
  total: number;
  started: number;
  now: number;
  complete?: boolean;
  error?: boolean;
};
const sendImportError = (label: string) =>
  window.dispatchEvent(
    new CustomEvent<ImportStatus>("billflow-import", {
      detail: {
        label,
        done: 0,
        total: 1,
        started: importStarted || Date.now(),
        now: Date.now(),
        complete: true,
        error: true,
      },
    }),
  );
let importStarted = 0;
const sendImportProgress = (
  label: string,
  done: number,
  total: number,
  complete = false,
) => {
  if (done === 0) importStarted = Date.now();
  window.dispatchEvent(
    new CustomEvent<ImportStatus>("billflow-import", {
      detail: {
        label,
        done,
        total,
        started: importStarted,
        now: Date.now(),
        complete,
      },
    }),
  );
};
const bulkImport = async (
  resource: string,
  inserts: Record<string, unknown>[],
  updates: Record<string, unknown>[],
  label: string,
) => {
  const jobs = [
      ...inserts.map((row) => ({ method: "POST", row })),
      ...updates.map((row) => ({ method: "PATCH", row })),
    ],
    total = jobs.length + 1;
  sendImportProgress(label, 0, total);
  let done = 0;
  for (let index = 0; index < jobs.length; index += 20) {
    const chunk = jobs.slice(index, index + 20);
    for (const method of ["POST", "PATCH"]) {
      const rows = chunk
        .filter((job) => job.method === method)
        .map((job) => job.row);
      if (!rows.length) continue;
      const controller = new AbortController(),
        timeout = setTimeout(() => controller.abort(), 30000);
      let response: Response;
      try {
        response = await fetch(`/api/data?resource=${resource}`, {
          method,
          headers: { "content-type": "application/json" },
          body: JSON.stringify(rows),
          signal: controller.signal,
        });
      } catch (error) {
        if (error instanceof DOMException && error.name === "AbortError")
          throw new Error(
            "Import batch timed out after 30 seconds. Clear the incomplete module data, then retry the file.",
          );
        throw error;
      } finally {
        clearTimeout(timeout);
      }
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Batch import failed");
    }
    done += chunk.length;
    sendImportProgress(label, done, total);
  }
  sendImportProgress(`${label} · linking related records`, done, total);
  const linkResponse = await fetch("/api/reconcile", { method: "POST" }),
    linkResult = await linkResponse.json();
  if (!linkResponse.ok)
    throw new Error(linkResult.error || "Automatic data linking failed");
  resourceCache.clear();
  window.dispatchEvent(new Event("billflow-data-changed"));
  sendImportProgress(label, total, total, true);
  return linkResult as Record<string, number>;
};
const ImportProgressCard = ({ progress }: { progress: ImportStatus }) => {
  const percent = progress.total
      ? Math.round((progress.done / progress.total) * 100)
      : 0,
    elapsed = Math.max(1, Math.round((progress.now - progress.started) / 1000)),
    remaining = progress.done
      ? Math.max(
          0,
          Math.round(
            (elapsed / progress.done) * (progress.total - progress.done),
          ),
        )
      : 0;
  return (
    <aside className={`import-progress ${progress.error ? "failed" : ""}`}>
      <div>
        <b>
          {progress.error
            ? "Import stopped"
            : progress.complete
              ? "Import complete"
              : "Importing data…"}
        </b>
        <span>{progress.label}</span>
      </div>
      <strong>{progress.error ? "!" : `${percent}%`}</strong>
      <div className="progress-track">
        <i style={{ width: `${percent}%` }} />
      </div>
      <small>
        {progress.error
          ? "Check the error message and file format."
          : `${progress.done} of ${progress.total} rows processed · ${progress.complete ? `${elapsed}s completed` : progress.done ? `about ${remaining}s remaining` : "calculating time…"}`}
      </small>
    </aside>
  );
};

export default function Home() {
  const [view, setView] = useState<View>("Dashboard");
  const [mobile, setMobile] = useState(false);
  const [query, setQuery] = useState("");
  const [toast, setToast] = useState("");
  const [summary, setSummary] = useState<Record<string, number>>({});
  const [importStatus, setImportStatus] = useState<ImportStatus | null>(null);
  const invoiceData = useResource("invoices"),
    productData = useResource("products");
  useEffect(() => {
    const refreshDashboardData = () => {
      fetch("/api/summary")
        .then((r) => r.json())
        .then(setSummary)
        .catch(() => {});
      invoiceData.reload();
      productData.reload();
    };
    refreshDashboardData();
    if (!sessionStorage.getItem("billflow-auto-linked-v1"))
      fetch("/api/reconcile", { method: "POST" })
        .then(async (response) => {
          if (!response.ok) return;
          sessionStorage.setItem("billflow-auto-linked-v1", "1");
          resourceCache.clear();
          refreshDashboardData();
        })
        .catch(() => {});
    for (const resource of [
      "customers",
      "suppliers",
      "products",
      "invoices",
      "purchases",
      "returns",
      "payments",
      "accounts",
      "transactions",
      "users",
      "settings",
    ])
      request(resource)
        .then((data) => resourceCache.set(resource, data.rows || []))
        .catch(() => {});
    window.addEventListener("billflow-data-changed", refreshDashboardData);
    return () =>
      window.removeEventListener("billflow-data-changed", refreshDashboardData);
  }, []);
  useEffect(() => {
    const update = (event: Event) => {
      const detail = (event as CustomEvent<ImportStatus>).detail;
      setImportStatus(detail);
      if (detail.complete) setTimeout(() => setImportStatus(null), 2500);
    };
    window.addEventListener("billflow-import", update);
    return () => window.removeEventListener("billflow-import", update);
  }, []);
  const invoices = useMemo(
    () =>
      invoiceData.rows.filter(
        (x) =>
          x.kind !== "quotation" &&
          JSON.stringify(x).toLowerCase().includes(query.toLowerCase()),
      ),
    [invoiceData.rows, query],
  );
  const go = (v: View) => {
    setView(v);
    setMobile(false);
  };
  const notify = (s: string) => {
    setToast(s);
    setTimeout(() => setToast(""), 2400);
  };
  return (
    <div className="app-shell">
      <aside className={mobile ? "sidebar open" : "sidebar"}>
        <div className="brand">
          <span className="brandmark">≋</span>
          <span>BillFlow</span>
        </div>
        <nav>
          {nav.map(([name, icon]) => (
            <button
              key={name}
              onClick={() => go(name)}
              className={view === name ? "active" : ""}
            >
              <span>{icon}</span>
              {name}
            </button>
          ))}
        </nav>
        <div className="business">
          <span className="store">▣</span>
          <div>
            <b>BillFlow Workspace</b>
            <small>Configure business profile</small>
          </div>
          <button onClick={() => go("Settings")}>›</button>
        </div>
      </aside>
      {mobile && (
        <button
          className="scrim"
          aria-label="Close menu"
          onClick={() => setMobile(false)}
        />
      )}
      <main>
        <header>
          <button className="menu" onClick={() => setMobile(true)}>
            ☰
          </button>
          <label className="search">
            <span>⌕</span>
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search invoices, customers, products..."
            />
          </label>
          <div className="profile">
            <span className="avatar">BF</span>
            <div>
              <b>BillFlow</b>
              <small>Live business workspace</small>
            </div>
          </div>
        </header>
        <div className="content">
          <div className="page-head">
            <div>
              <p>BUSINESS OVERVIEW</p>
              <h1>{view}</h1>
            </div>
            <div className="actions">
              <button
                className="secondary"
                onClick={() => {
                  go("Quotations");
                  notify("New quotation form opened");
                }}
              >
                + Quotation
              </button>
              <button
                className="primary"
                onClick={() => {
                  go("Invoices");
                  notify("New invoice form opened");
                }}
              >
                + Create Invoice
              </button>
            </div>
          </div>
          {view === "Dashboard" ? (
            <Dashboard
              go={go}
              summary={summary}
              invoices={invoices}
              products={productData.rows}
            />
          ) : (
            <Section view={view} invoices={invoices} notify={notify} />
          )}
        </div>
      </main>
      {importStatus && <ImportProgressCard progress={importStatus} />}
      {toast && <div className="toast">✓ {toast}</div>}
    </div>
  );
}

function Dashboard({
  go,
  summary,
  invoices,
  products,
}: {
  go: (v: View) => void;
  summary: Record<string, number>;
  invoices: DataRow[];
  products: DataRow[];
}) {
  return (
    <>
      <section className="stats">
        <Stat
          icon="↗"
          tone="blue"
          label="Total Sales"
          value={money(summary.sales || 0)}
          note="From saved invoices"
        />
        <Stat
          icon="₹"
          tone="amber"
          label="Total Purchases"
          value={money(summary.purchases || 0)}
          note="From imported purchases"
        />
        <Stat
          icon="▣"
          tone="green"
          label="Payments Collected"
          value={money(summary.payments || 0)}
          note="All payment entries"
        />
        <Stat
          icon="◇"
          tone="red"
          label="Low Stock"
          value={`${summary.lowStock || 0} items`}
          note="View inventory →"
          onClick={() => go("Inventory")}
        />
      </section>
      <DashboardRankings />
      <section className="money-strip">
        <button onClick={() => go("Cash Account")}>
          <span>Net Account Movement</span>
          <b>{money(summary.cash || 0)}</b>
          <small>Cash and bank transactions</small>
        </button>
        <button onClick={() => go("Invoices")}>
          <span>Sales Invoices</span>
          <b>{invoices.length}</b>
          <small>Saved records</small>
        </button>
        <button onClick={() => go("Purchases")}>
          <span>Total Purchases</span>
          <b>{money(summary.purchases || 0)}</b>
          <small>Imported and entered data</small>
        </button>
        <button onClick={() => go("Sales Returns")}>
          <span>Sales Returns</span>
          <b>{money(summary.returns || 0)}</b>
          <small>Saved credit notes</small>
        </button>
      </section>
      <section className="analytics">
        <article className="card chart-card">
          <div className="card-title">
            <h2>Sales & Purchase Overview</h2>
            <button onClick={() => go("Reports")}>View full report</button>
          </div>
          {summary.sales || summary.purchases ? (
            <div className="live-bars">
              <div>
                <span
                  style={{
                    height: `${Math.max(6, ((summary.sales || 0) / Math.max(summary.sales || 0, summary.purchases || 0)) * 100)}%`,
                  }}
                />
                <b>Sales</b>
                <small>{money(summary.sales || 0)}</small>
              </div>
              <div>
                <span
                  className="purchase"
                  style={{
                    height: `${Math.max(6, ((summary.purchases || 0) / Math.max(summary.sales || 0, summary.purchases || 0)) * 100)}%`,
                  }}
                />
                <b>Purchases</b>
                <small>{money(summary.purchases || 0)}</small>
              </div>
              <div>
                <span
                  className="returns"
                  style={{
                    height: `${Math.max(6, ((summary.returns || 0) / Math.max(summary.sales || 0, summary.purchases || 0)) * 100)}%`,
                  }}
                />
                <b>Returns</b>
                <small>{money(summary.returns || 0)}</small>
              </div>
            </div>
          ) : (
            <Empty text="Add invoices or import purchases to generate this chart." />
          )}
        </article>
        <article className="card collection">
          <div className="card-title">
            <h2>Live Business Data</h2>
            <button onClick={() => go("Reports")}>Open Reports</button>
          </div>
          <div className="donut-row">
            <div className="donut">
              <div>
                <b>{money(summary.sales || 0)}</b>
                <span>Total Sales</span>
              </div>
            </div>
            <ul>
              <li>
                <i className="blue-dot" />
                Sales <b>{money(summary.sales || 0)}</b>
              </li>
              <li>
                <i className="green-dot" />
                Payments <b>{money(summary.payments || 0)}</b>
              </li>
              <li>
                <i className="amber-dot" />
                Purchases <b>{money(summary.purchases || 0)}</b>
              </li>
            </ul>
          </div>
          <div className="info">
            ⓘ Figures update from your saved database records
          </div>
        </article>
      </section>
      <section className="tables">
        <TableCard title="Recent Invoices" action={() => go("Invoices")}>
          <InvoiceTable rows={invoices.slice(0, 5)} />
        </TableCard>
        <TableCard title="Stock Alerts" action={() => go("Inventory")}>
          <table>
            <thead>
              <tr>
                <th>Product</th>
                <th>Remaining</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {products
                .filter((p) => Number(p.stock) <= Number(p.reorder_level))
                .slice(0, 5)
                .map((p) => (
                  <tr key={p.id}>
                    <td>
                      <b>{String(p.name)}</b>
                      <small>{String(p.category || "")}</small>
                    </td>
                    <td className={Number(p.stock) === 0 ? "danger" : ""}>
                      {String(p.stock)} units
                    </td>
                    <td>
                      <span
                        className={
                          Number(p.stock) === 0
                            ? "pill overdue"
                            : "pill pending"
                        }
                      >
                        {Number(p.stock) === 0 ? "Out of Stock" : "Low Stock"}
                      </span>
                    </td>
                  </tr>
                ))}
            </tbody>
          </table>
        </TableCard>
      </section>
    </>
  );
}
function DashboardRankings() {
  const [rankings, setRankings] = useState<{
    bestSellers: Record<string, unknown>[];
    bestDealers: Record<string, unknown>[];
  }>({ bestSellers: [], bestDealers: [] });
  useEffect(() => {
    fetch("/api/audit")
      .then((r) => r.json())
      .then((data) =>
        setRankings({
          bestSellers: data.bestSellers || [],
          bestDealers: data.bestDealers || [],
        }),
      )
      .catch(() => {});
  }, []);
  return (
    <section className="dashboard-rankings">
      <article className="card best-sellers">
        <div className="card-title">
          <h2>Best-Selling Products</h2>
          <small>Ranked by quantity sold</small>
        </div>
        {rankings.bestSellers.length ? (
          <div className="best-grid">
            {rankings.bestSellers.slice(0, 5).map((x, i) => (
              <article key={`${x.code}-${x.item}`}>
                <span>#{i + 1}</span>
                <div>
                  <b>{String(x.item)}</b>
                  <small>{String(x.code || "No item code")}</small>
                </div>
                <strong>
                  {Number(x.quantity).toLocaleString("en-IN")} sold ·{" "}
                  {money(Number(x.revenue))}
                </strong>
              </article>
            ))}
          </div>
        ) : (
          <Empty text="Import sales data to calculate best-selling products." />
        )}
      </article>
      <article className="card best-sellers best-dealers">
        <div className="card-title">
          <h2>Best Dealers</h2>
          <small>Ranked by sales purchase value</small>
        </div>
        {rankings.bestDealers.length ? (
          <div className="best-grid">
            {rankings.bestDealers.slice(0, 5).map((x, i) => (
              <article key={`${x.code}-${x.name}`}>
                <span>#{i + 1}</span>
                <div>
                  <b>{String(x.name)}</b>
                  <small>{String(x.code || "No dealer code")}</small>
                </div>
                <strong>
                  {money(Number(x.purchaseValue))} ·{" "}
                  {Number(x.bills).toLocaleString("en-IN")} bills
                </strong>
              </article>
            ))}
          </div>
        ) : (
          <Empty text="Import sales data to calculate the best dealers." />
        )}
      </article>
    </section>
  );
}

function Stat({
  icon,
  tone,
  label,
  value,
  note,
  onClick,
}: {
  icon: string;
  tone: string;
  label: string;
  value: string;
  note: string;
  onClick?: () => void;
}) {
  return (
    <button className="stat" onClick={onClick}>
      <span className={`stat-icon ${tone}`}>{icon}</span>
      <div>
        <span>{label}</span>
        <strong>{value}</strong>
        <small>{note}</small>
      </div>
    </button>
  );
}
function TableCard({
  title,
  action,
  children,
}: {
  title: string;
  action: () => void;
  children: React.ReactNode;
}) {
  return (
    <article className="card table-card">
      <div className="card-title">
        <h2>{title}</h2>
        <button className="link" onClick={action}>
          View All ›
        </button>
      </div>
      {children}
    </article>
  );
}
function InvoiceTable({ rows }: { rows: DataRow[] }) {
  return rows.length ? (
    <table>
      <thead>
        <tr>
          <th>Invoice</th>
          <th>Customer</th>
          <th>Date</th>
          <th>Amount</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        {rows.map((x) => (
          <tr key={x.id}>
            <td>
              <b>{String(x.invoice_no || x.id)}</b>
            </td>
            <td>{String(x.customer_name || "")}</td>
            <td>{formatDate(x.created_at)}</td>
            <td>
              <b>{money(Number(x.amount) || 0)}</b>
            </td>
            <td>
              <span
                className={`pill ${String(x.status || "pending").toLowerCase()}`}
              >
                {String(x.status || "Pending")}
              </span>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  ) : (
    <Empty text="No invoices yet. Create your first sales invoice." />
  );
}
function Section({
  view,
  invoices,
  notify,
}: {
  view: View;
  invoices: DataRow[];
  notify: (s: string) => void;
}) {
  const [showForm, setShowForm] = useState(false);
  return (
    <section className="card section-page">
      <div className="empty-hero">
        <span>{nav.find((n) => n[0] === view)?.[1]}</span>
        <div>
          <h2>{view} Management</h2>
          <p>{sectionDescription(view)}</p>
        </div>
      </div>
      {view === "Invoices" ? (
        <SalesInvoicePanel rows={invoices} notify={notify} />
      ) : view === "Sales Returns" ? (
        <ReturnsPanel notify={notify} />
      ) : view === "Purchases" ? (
        <PurchasesPanel notify={notify} />
      ) : view === "Cash Account" ? (
        <AccountsPanel mode="cash" notify={notify} />
      ) : view === "Bank Accounts" ? (
        <AccountsPanel mode="bank" notify={notify} />
      ) : view === "Data Errors" ? (
        <DataErrorsPanel notify={notify} />
      ) : view === "Users" ? (
        <UsersPanel notify={notify} />
      ) : view === "Inventory" ? (
        <InventoryPanel notify={notify} />
      ) : view === "Payments" ? (
        <PaymentsPanel notify={notify} />
      ) : view === "Quotations" ? (
        <QuotationPanel notify={notify} />
      ) : view === "Customers" ? (
        <CustomersPanel notify={notify} />
      ) : view === "Suppliers" ? (
        <SuppliersPanel notify={notify} />
      ) : view === "Reports" ? (
        <ReportsPanel />
      ) : view === "Settings" ? (
        <SettingsPanel notify={notify} />
      ) : (
        <div />
      )}
    </section>
  );
}

function sectionDescription(v: View) {
  return {
    Invoices: "Create, edit, import and export item-wise GST sales invoices.",
    "Sales Returns":
      "Create, import and export credit notes and sales returns.",
    Purchases: "Import and export purchase invoices through Excel.",
    "Cash Account":
      "Maintain a separate cash book with Excel import and export.",
    "Bank Accounts":
      "Maintain bank accounts and transactions with Excel import and export.",
    "Data Errors":
      "Automatically find mismatches across sales, purchases, masters, stock, cash and bank data.",
    Quotations: "Prepare quotations and convert approved quotes into invoices.",
    Customers: "Import and maintain complete customer master data from Excel.",
    Suppliers: "Import and maintain complete supplier master data from Excel.",
    Inventory: "Track purchases, sales stock and low-stock levels.",
    Payments: "Record collections and automatically update customer balances.",
    Reports: "Visualise sales, purchases, returns, profit, cash flow and GST.",
    Users: "Control staff access, roles and account authorization.",
    Settings: "Configure your business profile, taxes and invoice numbering.",
    Dashboard: "",
  }[v];
}
function QuickForm({
  view,
  close,
  notify,
}: {
  view: View;
  close: () => void;
  notify: (s: string) => void;
}) {
  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    notify(`${view.slice(0, -1)} saved successfully`);
    close();
  };
  return (
    <div className="form-wrap">
      <form onSubmit={submit}>
        <div className="form-head">
          <div>
            <h3>New {view.slice(0, -1)}</h3>
            <p>Enter the details below. Fields marked * are required.</p>
          </div>
          <button type="button" onClick={close}>
            ×
          </button>
        </div>
        <div className="form-grid">
          <label>
            Customer / Party *
            <input required placeholder="Select or type customer name" />
          </label>
          <label>
            Date *<input required type="date" defaultValue="2026-07-15" />
          </label>
          {view === "Inventory" ? (
            <>
              <label>
                Product name *<input required placeholder="Product name" />
              </label>
              <label>
                Opening quantity *
                <input required type="number" min="0" placeholder="0" />
              </label>
            </>
          ) : (
            <>
              <label>
                Reference number
                <input placeholder="Auto generated" />
              </label>
              <label>
                Amount (₹) *
                <input required type="number" min="1" placeholder="0.00" />
              </label>
            </>
          )}
        </div>
        <label className="wide">
          Notes
          <textarea placeholder="Optional notes or payment terms" />
        </label>
        <div className="form-actions">
          <button type="button" className="secondary" onClick={close}>
            Cancel
          </button>
          <button className="primary">Save {view.slice(0, -1)}</button>
        </div>
      </form>
    </div>
  );
}
function SalesInvoicePanel({
  rows,
  notify,
}: {
  rows: DataRow[];
  notify: (s: string) => void;
}) {
  const products = useResource("products"),
    customers = useResource("customers"),
    [open, setOpen] = useState(false),
    [editing, setEditing] = useState<DataRow | null>(null),
    [viewing, setViewing] = useState<DataRow | null>(null),
    [busy, setBusy] = useState(false),
    [limit, setLimit] = useState(25),
    [filter, setFilter] = useState("All"),
    [query, setQuery] = useState(""),
    [sort, setSort] = useState<"billAsc" | "billDesc" | "newest" | "oldest">(
      "billAsc",
    );
  const visibleRows = rows
    .filter(
      (row) =>
        (filter === "All" ||
          String(row.status).toLowerCase() === filter.toLowerCase()) &&
        includesQuery(row, query),
    )
    .sort((a, b) =>
      sort === "newest" || sort === "oldest"
        ? String(a.created_at).localeCompare(String(b.created_at)) *
          (sort === "oldest" ? 1 : -1)
        : billCompare(
            a.invoice_no,
            b.invoice_no,
            sort === "billAsc" ? "asc" : "desc",
          ),
    );
  const enrichSalesItem = (item: Record<string, unknown>) => {
    const product = findProduct(
        products.rows,
        item.code,
        item.item || item.name || item.description,
      ),
      quantity = Number(item.quantity || item.qty || 0),
      rate = Number(item.rate || item.unit_price || product?.price || 0),
      total = Number(item.total || item.product_value || quantity * rate || 0);
    return {
      ...item,
      code: item.code || product?.sku || "",
      item:
        item.item ||
        item.name ||
        item.description ||
        product?.name ||
        `Product ${String(item.code || "")}`.trim(),
      quantity,
      rate,
      total,
    };
  };
  const invoiceItems = (row: DataRow) => {
    try {
      const items = JSON.parse(String(row.items_json || "[]"));
      if (Array.isArray(items) && items.length)
        return items.map(enrichSalesItem);
    } catch {}
    return [
      {
        item: "Invoice total",
        quantity: 1,
        rate: row.amount,
        total: row.amount,
      },
    ];
  };
  const save = async (body: Record<string, unknown>) => {
    await request(
      "invoices",
      editing ? "PATCH" : "POST",
      editing ? { ...body, id: editing.id } : body,
    );
    notify(editing ? "Invoice updated" : "Invoice saved");
    location.reload();
  };
  const remove = async (id: number) => {
    if (!confirm("Delete this invoice?")) return;
    await request("invoices", "DELETE", undefined, id);
    notify("Invoice deleted");
    location.reload();
  };
  const importFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setBusy(true);
    try {
      const list = await readExcel(f);
      const grouped = new Map<string, Record<string, unknown>[]>();
      let lastBill = "";
      for (const [rowIndex, row] of list.entries()) {
        const enteredBill = String(cell(row, "bill no.", "bill no")).trim(),
          bill = enteredBill || lastBill || `IMPORT-${Date.now()}-${rowIndex}`;
        if (enteredBill) lastBill = enteredBill;
        grouped.set(bill, [...(grouped.get(bill) || []), row]);
      }
      const inserts: Record<string, unknown>[] = [],
        updates: Record<string, unknown>[] = [];
      let duplicateLines = 0,
        duplicateBills = 0,
        importedLines = 0;
      for (const [i, [bill, lines]] of [...grouped.entries()].entries()) {
        const first = lines[0],
          existing = rows.find((x) => String(x.invoice_no) === bill),
          customerCode = String(
            lines
              .map((r) => cell(r, "customer code"))
              .find((v) => String(v).trim()) || "",
          ).trim(),
          matchedCustomer = customers.rows.find(
            (customer) =>
              customerCode &&
              normalizedValue(customer.code) === normalizedValue(customerCode),
          ),
          customerName = String(
            lines
              .map((r) => cell(r, "customer name"))
              .find((v) => String(v).trim()) ||
              matchedCustomer?.name ||
              existing?.customer_name ||
              customerCode,
          ).trim(),
          uniqueIncoming = new Map<string, Record<string, unknown>>();
        for (const line of lines) {
          const code = String(cell(line, "code")).trim(),
            enteredName = String(cell(line, "item")).trim(),
            product = findProduct(products.rows, code, enteredName),
            quantity = excelNumber(cell(line, "qty")),
            enteredTotal = excelNumber(cell(line, "product value")),
            enteredRate = excelNumber(cell(line, "unit price")),
            valueRate = quantity && enteredTotal ? enteredTotal / quantity : 0,
            rate =
              enteredRate ||
              valueRate ||
              Number(product?.price || product?.purchase_rate || 0),
            total = enteredTotal || quantity * rate,
            item = {
              code: code || String(product?.sku || ""),
              item:
                enteredName ||
                String(product?.name || "") ||
                `Product ${code}`.trim(),
              quantity,
              rate,
              total,
              rate_source: enteredRate
                ? "sales"
                : valueRate
                  ? "sales-value"
                  : product?.price
                    ? "stock"
                    : product?.purchase_rate
                      ? "purchase"
                      : "",
            },
            key = rowFingerprint(
              item.code || item.item,
              item.quantity,
              item.rate,
              item.total,
            );
          if (uniqueIncoming.has(key)) duplicateLines += 1;
          else uniqueIncoming.set(key, item);
        }
        const incomingItems = [...uniqueIncoming.values()],
          existingItems = existing ? invoiceItems(existing) : [],
          existingKeys = new Set(
            existingItems.map((item) =>
              rowFingerprint(
                item.code || item.item,
                item.quantity,
                item.rate,
                item.total,
              ),
            ),
          ),
          newItems = incomingItems.filter((item) => {
            const key = rowFingerprint(
              item.code || item.item,
              item.quantity,
              item.rate,
              item.total,
            );
            if (existingKeys.has(key)) {
              duplicateLines += 1;
              return false;
            }
            return true;
          });
        if (existing && !newItems.length) {
          duplicateBills += 1;
          continue;
        }
        const items = existing
            ? [...existingItems, ...newItems]
            : incomingItems,
          subtotal = items.reduce(
            (sum, item) => sum + Number(item.total || 0),
            0,
          ),
          enteredDate = cell(first, "bill date");
        importedLines += existing ? newItems.length : incomingItems.length;
        const body = {
          invoice_no: bill || `INV-IMPORT-${Date.now()}-${i}`,
          customer_name: customerName || "Unknown Customer",
          customer_code: customerCode,
          kind: "invoice",
          items_json: JSON.stringify(items),
          subtotal,
          tax: 0,
          amount: subtotal,
          status: existing?.status || "Pending",
          created_at: String(enteredDate).trim()
            ? excelDate(enteredDate)
            : String(
                existing?.created_at || new Date().toISOString().slice(0, 10),
              ),
        };
        if (existing) updates.push({ ...body, id: existing.id });
        else inserts.push(body);
      }
      await bulkImport("invoices", inserts, updates, "Sales report");
      notify(
        `${inserts.length + updates.length} sales bills updated · ${importedLines} new item rows${duplicateLines ? ` · ${duplicateLines} duplicates skipped` : ""}${duplicateBills ? ` · ${duplicateBills} unchanged bills skipped` : ""}`,
      );
      location.reload();
    } catch (err) {
      sendImportError("Sales report");
      notify(err instanceof Error ? err.message : "Excel import failed");
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };
  const exportRows = () =>
    exportExcel(
      rows.flatMap((x) => {
        const items = invoiceItems(x);
        return (items.length ? items : [{}]).map((item) => ({
          "Customer Code": x.customer_code || "",
          "Customer Name": x.customer_name,
          "Bill No.": x.invoice_no,
          "Bill Date": formatDate(x.created_at),
          Code: item.code || "",
          Item: item.item || item.description || "",
          Qty: item.quantity || 0,
          "Unit Price": item.rate || 0,
          "Product Value": item.total || x.amount,
        }));
      }),
      "sales-invoices",
    );
  return (
    <div>
      <div className="panel-tools">
        <div className="filter-tabs">
          {["All", "Paid", "Pending", "Overdue"].map((value) => (
            <button
              key={value}
              className={filter === value ? "chosen" : ""}
              onClick={() => setFilter(value)}
            >
              {value === "All" ? "All invoices" : value}
            </button>
          ))}
        </div>
        <div className="excel-actions">
          <label className="secondary upload">
            {busy ? "Importing..." : "Import Excel"}
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={importFile}
              disabled={busy}
            />
          </label>
          <button className="secondary" onClick={exportRows}>
            Export Excel
          </button>
          <button
            className="primary"
            onClick={() => {
              setEditing(null);
              setOpen(!open);
            }}
          >
            {open ? "Close Form" : "+ Sales Invoice"}
          </button>
        </div>
      </div>
      <div className="smart-import-note">
        ✓ Product codes link Sales, Stock and Purchases automatically. Missing
        unit prices use Product Value ÷ Qty first, then the linked saved rate.
        Repeated bill lines are skipped.
      </div>
      {open && (
        <InvoiceBuilder
          initial={editing}
          save={save}
          close={() => setOpen(false)}
        />
      )}{" "}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search bill number, customer, product or code"
      >
        <label>
          Sort invoices
          <select
            value={sort}
            onChange={(event) =>
              setSort(
                event.target.value as
                  | "billAsc"
                  | "billDesc"
                  | "newest"
                  | "oldest",
              )
            }
          >
            <option value="billAsc">Bill No. A–Z</option>
            <option value="billDesc">Bill No. Z–A</option>
            <option value="newest">Date: new to old</option>
            <option value="oldest">Date: old to new</option>
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={visibleRows.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (!confirm("Clear all sales invoices?")) return;
          await clearResource("invoices", "&kind=invoice");
          notify("Sales data cleared");
          location.reload();
        }}
      />
      {visibleRows.length ? (
        <div className="excel-preview">
          <table>
            <thead>
              <tr>
                <th>Invoice</th>
                <th>Customer</th>
                <th>Date</th>
                <th>Amount</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {visibleRows.slice(0, limit).map((x) => (
                <tr key={x.id}>
                  <td>
                    <b>{String(x.invoice_no)}</b>
                  </td>
                  <td>{String(x.customer_name)}</td>
                  <td>{formatDate(x.created_at)}</td>
                  <td>{money(Number(x.amount))}</td>
                  <td>
                    <span className={`pill ${String(x.status).toLowerCase()}`}>
                      {String(x.status)}
                    </span>
                  </td>
                  <td className="row-actions">
                    <button onClick={() => setViewing(x)}>View</button>
                    <button
                      onClick={() => {
                        setEditing(x);
                        setOpen(true);
                      }}
                    >
                      Edit
                    </button>
                    <button className="danger" onClick={() => remove(x.id)}>
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <Empty text="No sales invoices yet. Import Excel or create your first invoice." />
      )}
      {viewing && (
        <BillDetails
          title={`Sales Bill ${String(viewing.invoice_no)}`}
          party={String(viewing.customer_name || "Unknown customer")}
          date={String(viewing.created_at || "")}
          items={invoiceItems(viewing)}
          close={() => setViewing(null)}
        />
      )}
    </div>
  );
}
function InvoiceBuilder({
  initial,
  save,
  close,
  kind = "invoice",
}: {
  initial: DataRow | null;
  save: (b: Record<string, unknown>) => Promise<void>;
  close: () => void;
  kind?: "invoice" | "quotation";
}) {
  const parsedItems = initial?.items_json
      ? JSON.parse(String(initial.items_json))
      : null,
    oldItems = parsedItems?.map((x: Record<string, unknown>) => ({
      name: String(x.name || x.item || x.description || ""),
      qty: Number(x.qty || x.quantity || 1),
      rate: Number(x.rate || 0),
      gst: Number(x.gst || 0),
    }));
  const [customer, setCustomer] = useState(
      String(initial?.customer_name || ""),
    ),
    [date, setDate] = useState(
      String(initial?.created_at || new Date().toISOString().slice(0, 10)),
    ),
    [status, setStatus] = useState(String(initial?.status || "Pending")),
    [items, setItems] = useState(
      oldItems?.length ? oldItems : [{ name: "", qty: 1, rate: 0, gst: 18 }],
    ),
    [busy, setBusy] = useState(false);
  const subtotal = items.reduce(
      (s: number, x: { qty: number; rate: number }) =>
        s + Number(x.qty) * Number(x.rate),
      0,
    ),
    tax = items.reduce(
      (s: number, x: { qty: number; rate: number; gst: number }) =>
        s + (Number(x.qty) * Number(x.rate) * Number(x.gst)) / 100,
      0,
    ),
    total = subtotal + tax;
  return (
    <form
      className="invoice-builder"
      onSubmit={async (e) => {
        e.preventDefault();
        setBusy(true);
        try {
          await save({
            invoice_no: String(
              initial?.invoice_no ||
                `${kind === "quotation" ? "QT" : "INV"}-${Date.now().toString().slice(-8)}`,
            ),
            customer_name: customer,
            kind,
            items_json: JSON.stringify(items),
            subtotal,
            tax,
            amount: total,
            status,
            created_at: date,
          });
        } finally {
          setBusy(false);
        }
      }}
    >
      <div className="form-grid">
        <label>
          Customer *
          <input
            value={customer}
            onChange={(e) => setCustomer(e.target.value)}
            required
          />
        </label>
        <label>
          Invoice date *
          <input
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            required
          />
        </label>
        <label>
          Status
          <select value={status} onChange={(e) => setStatus(e.target.value)}>
            <option>Pending</option>
            <option>Paid</option>
            <option>Overdue</option>
          </select>
        </label>
        <label>
          Invoice Number
          <input
            value={String(initial?.invoice_no || "Auto-generated")}
            disabled
          />
        </label>
      </div>
      <div className="line-items">
        <div className="line-head">
          <b>Item</b>
          <b>Qty</b>
          <b>Rate</b>
          <b>GST</b>
          <b>Amount</b>
          <span />
        </div>
        {items.map(
          (
            x: { name: string; qty: number; rate: number; gst: number },
            i: number,
          ) => (
            <div className="line-row" key={i}>
              <input
                required
                value={x.name}
                onChange={(e) =>
                  setItems((a: typeof items) =>
                    a.map((v: typeof x, j: number) =>
                      j === i ? { ...v, name: e.target.value } : v,
                    ),
                  )
                }
              />
              <input
                type="number"
                min="0.01"
                step=".01"
                value={x.qty}
                onChange={(e) =>
                  setItems((a: typeof items) =>
                    a.map((v: typeof x, j: number) =>
                      j === i ? { ...v, qty: +e.target.value } : v,
                    ),
                  )
                }
              />
              <input
                type="number"
                min="0"
                step=".01"
                value={x.rate}
                onChange={(e) =>
                  setItems((a: typeof items) =>
                    a.map((v: typeof x, j: number) =>
                      j === i ? { ...v, rate: +e.target.value } : v,
                    ),
                  )
                }
              />
              <select
                value={x.gst}
                onChange={(e) =>
                  setItems((a: typeof items) =>
                    a.map((v: typeof x, j: number) =>
                      j === i ? { ...v, gst: +e.target.value } : v,
                    ),
                  )
                }
              >
                <option>18</option>
                <option>12</option>
                <option>5</option>
                <option>0</option>
              </select>
              <b>{money(x.qty * x.rate)}</b>
              <button
                type="button"
                onClick={() =>
                  setItems((a: typeof items) =>
                    a.filter((_: unknown, j: number) => j !== i),
                  )
                }
              >
                ×
              </button>
            </div>
          ),
        )}
        <button
          type="button"
          className="add-line"
          onClick={() =>
            setItems((a: typeof items) => [
              ...a,
              { name: "", qty: 1, rate: 0, gst: 18 },
            ])
          }
        >
          + Add item
        </button>
      </div>
      <div className="invoice-total">
        <span>
          Subtotal <b>{money(subtotal)}</b>
        </span>
        <span>
          GST <b>{money(tax)}</b>
        </span>
        <strong>
          Grand Total <b>{money(total)}</b>
        </strong>
      </div>
      <div className="form-actions">
        <button type="button" className="secondary" onClick={close}>
          Cancel
        </button>
        <button className="primary" disabled={busy}>
          {busy
            ? "Saving..."
            : `Save ${kind === "quotation" ? "Quotation" : "Invoice"}`}
        </button>
      </div>
    </form>
  );
}
function ReturnsPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("returns"),
    invoices = useResource("invoices"),
    customers = useResource("customers"),
    [open, setOpen] = useState(false),
    [busy, setBusy] = useState(false),
    [query, setQuery] = useState(""),
    [sort, setSort] = useState<"newest" | "oldest">("newest"),
    [limit, setLimit] = useState(25);
  const visibleReturns = data.rows
    .filter((row) => {
      const invoice = invoices.rows.find(
          (item) => item.id === Number(row.invoice_id),
        ),
        customer = customers.rows.find(
          (item) => item.id === Number(row.customer_id),
        );
      return includesQuery(
        {
          ...row,
          invoice_no: invoice?.invoice_no || "",
          customer_name: customer?.name || "",
        },
        query,
      );
    })
    .sort(
      (a, b) =>
        String(a.created_at).localeCompare(String(b.created_at)) *
        (sort === "oldest" ? 1 : -1),
    );
  const save = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    await request("returns", "POST", {
      credit_note_no: `CN-${Date.now().toString().slice(-8)}`,
      invoice_id: Number(f.get("invoice_id")),
      customer_id: Number(f.get("customer_id")),
      amount: Number(f.get("amount")),
      reason: String(f.get("reason")),
      created_at: String(f.get("created_at")),
    });
    notify("Sales return saved");
    setOpen(false);
    data.reload();
  };
  const importFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setBusy(true);
    try {
      const list = await readExcel(f);
      const creditNotes = new Set(
          data.rows.map((row) => normalizedValue(row.credit_note_no)),
        ),
        fingerprints = new Set(
          data.rows.map((row) =>
            rowFingerprint(
              row.invoice_id,
              row.customer_id,
              row.amount,
              row.reason,
              excelDate(row.created_at),
            ),
          ),
        ),
        inserts: Record<string, unknown>[] = [];
      let duplicates = 0;
      for (const [i, row] of list.entries()) {
        const invoiceId = Number(cell(row, "invoice id", "invoice_id") || 0),
          customerId = Number(cell(row, "customer id", "customer_id") || 0),
          amount = excelNumber(cell(row, "amount", "return amount")),
          reason = String(cell(row, "reason", "description")),
          createdAt = excelDate(cell(row, "date", "return date")),
          enteredCreditNote = String(
            cell(row, "credit note", "credit note no", "credit_note_no"),
          ).trim(),
          fingerprint = rowFingerprint(
            invoiceId,
            customerId,
            amount,
            reason,
            createdAt,
          ),
          creditNote =
            enteredCreditNote ||
            `CN-${createdAt.replaceAll("-", "")}-${invoiceId || customerId}-${i + 1}`;
        if (
          creditNotes.has(normalizedValue(creditNote)) ||
          fingerprints.has(fingerprint)
        ) {
          duplicates += 1;
          continue;
        }
        creditNotes.add(normalizedValue(creditNote));
        fingerprints.add(fingerprint);
        inserts.push({
          credit_note_no: creditNote,
          invoice_id: invoiceId,
          customer_id: customerId,
          amount,
          reason,
          created_at: createdAt,
        });
      }
      await bulkImport("returns", inserts, [], "Sales returns");
      notify(
        `${inserts.length} sales returns imported${duplicates ? ` · ${duplicates} duplicates skipped` : ""}`,
      );
      data.reload();
    } catch (err) {
      sendImportError("Sales returns");
      notify(err instanceof Error ? err.message : "Excel import failed");
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };
  const remove = async (id: number) => {
    await request("returns", "DELETE", undefined, id);
    notify("Sales return deleted");
    data.reload();
  };
  return (
    <div className="panel-pad">
      <div className="return-summary">
        <article>
          <span>Total sales returns</span>
          <b>{money(data.rows.reduce((s, x) => s + Number(x.amount), 0))}</b>
        </article>
        <article>
          <span>Credit notes</span>
          <b>{data.rows.length}</b>
        </article>
        <article>
          <span>Database status</span>
          <b>Live</b>
        </article>
      </div>
      <div className="excel-actions">
        <label className="secondary upload">
          {busy ? "Importing..." : "Import Excel"}
          <input type="file" accept=".xlsx,.xls,.csv" onChange={importFile} />
        </label>
        <button
          className="secondary"
          onClick={() =>
            exportExcel(
              data.rows.map((x) => ({
                "Credit Note": x.credit_note_no,
                "Invoice ID": x.invoice_id,
                "Customer ID": x.customer_id,
                Amount: x.amount,
                Reason: x.reason,
                Date: formatDate(x.created_at),
              })),
              "sales-returns",
            )
          }
        >
          Export Excel
        </button>
        <button className="primary" onClick={() => setOpen(!open)}>
          + Create Sales Return
        </button>
      </div>
      {open && (
        <form className="invoice-builder" onSubmit={save}>
          <div className="form-grid">
            <label>
              Invoice *
              <select name="invoice_id" required>
                <option value="">Select invoice</option>
                {invoices.rows
                  .filter((x) => x.kind !== "quotation")
                  .map((x) => (
                    <option key={x.id} value={x.id}>
                      {String(x.invoice_no)} · {String(x.customer_name)}
                    </option>
                  ))}
              </select>
            </label>
            <label>
              Customer *
              <select name="customer_id" required>
                <option value="">Select customer</option>
                {customers.rows.map((x) => (
                  <option key={x.id} value={x.id}>
                    {String(x.code || "")} {String(x.name)}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Return amount *
              <input name="amount" type="number" step=".01" required />
            </label>
            <label>
              Date *
              <input
                name="created_at"
                type="date"
                defaultValue={new Date().toISOString().slice(0, 10)}
                required
              />
            </label>
          </div>
          <label className="wide">
            Reason
            <textarea name="reason" />
          </label>
          <div className="form-actions">
            <button className="primary">Save Credit Note</button>
          </div>
        </form>
      )}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search credit note, invoice, customer or reason"
      >
        <label>
          Date order
          <select
            value={sort}
            onChange={(event) =>
              setSort(event.target.value as "newest" | "oldest")
            }
          >
            <option value="newest">Newest first</option>
            <option value="oldest">Oldest first</option>
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={visibleReturns.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (!confirm("Clear all sales returns?")) return;
          await clearResource("returns");
          notify("Sales return data cleared");
          data.reload();
        }}
      />
      {visibleReturns.length ? (
        <div className="excel-preview">
          <table>
            <thead>
              <tr>
                <th>Credit Note</th>
                <th>Invoice</th>
                <th>Customer</th>
                <th>Amount</th>
                <th>Date</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {visibleReturns.slice(0, limit).map((x) => (
                <tr key={x.id}>
                  <td>{String(x.credit_note_no)}</td>
                  <td>
                    {String(
                      invoices.rows.find(
                        (invoice) => invoice.id === Number(x.invoice_id),
                      )?.invoice_no || `#${String(x.invoice_id)}`,
                    )}
                  </td>
                  <td>
                    {String(
                      customers.rows.find(
                        (customer) => customer.id === Number(x.customer_id),
                      )?.name || `#${String(x.customer_id)}`,
                    )}
                  </td>
                  <td>{money(Number(x.amount))}</td>
                  <td>{formatDate(x.created_at)}</td>
                  <td>
                    <button className="danger" onClick={() => remove(x.id)}>
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <Empty text="No sales returns saved. Import Excel or add one." />
      )}
    </div>
  );
}
function PurchasesPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("purchases"),
    products = useResource("products"),
    suppliers = useResource("suppliers"),
    [file, setFile] = useState(""),
    [busy, setBusy] = useState(false),
    [limit, setLimit] = useState(25),
    [query, setQuery] = useState(""),
    [sort, setSort] = useState<"billAsc" | "billDesc" | "newest" | "oldest">(
      "billAsc",
    ),
    [viewingBill, setViewingBill] = useState<
      (Record<string, unknown> & { items: DataRow[] }) | null
    >(null);
  const upload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setFile(f.name);
    setBusy(true);
    try {
      const allRows = await readExcel(f),
        rows = allRows.filter(
          (row) =>
            (String(cell(row, "supplier name")).trim() ||
              String(cell(row, "supplier code")).trim()) &&
            (String(cell(row, "item")).trim() ||
              String(cell(row, "code")).trim()) &&
            String(cell(row, "bill no.", "bill no")).trim(),
        ),
        incomplete = allRows.length - rows.length,
        seen = new Set(
          data.rows.map((row) =>
            rowFingerprint(
              row.bill_no,
              row.supplier_code || row.supplier,
              row.item_code || row.item,
              row.quantity,
              row.total,
              excelDate(row.purchase_date),
            ),
          ),
        );
      const inserts: Record<string, unknown>[] = [];
      let duplicates = 0;
      for (const r of rows) {
        const get = (...names: string[]) => cell(r, ...names),
          billNo = String(get("bill no.", "bill no")).trim(),
          supplierCode = String(get("supplier code")).trim(),
          enteredSupplier = String(get("supplier name")).trim(),
          supplierMaster = suppliers.rows.find(
            (supplier) =>
              (supplierCode &&
                normalizedValue(supplier.code) ===
                  normalizedValue(supplierCode)) ||
              (!supplierCode &&
                enteredSupplier &&
                normalizedValue(supplier.name) ===
                  normalizedValue(enteredSupplier)),
          ),
          itemCode = String(get("code")).trim(),
          enteredItem = String(get("item")).trim(),
          product = findProduct(products.rows, itemCode, enteredItem),
          quantity = excelNumber(get("qty")),
          enteredValue = excelNumber(get("product value")),
          rate =
            (quantity && enteredValue ? enteredValue / quantity : 0) ||
            Number(product?.purchase_rate || 0),
          total = enteredValue || quantity * rate,
          purchaseDate = excelDate(get("bill date")),
          receivedValue = get("rec.date", "rec. date"),
          body = {
            bill_no: billNo,
            supplier: String(
              enteredSupplier || supplierMaster?.name || supplierCode,
            ),
            supplier_code: String(supplierCode || supplierMaster?.code || ""),
            gstin: "",
            item: String(
              enteredItem || product?.name || `Product ${itemCode}`.trim(),
            ),
            item_code: String(itemCode || product?.sku || ""),
            quantity,
            rate,
            gst_amount: 0,
            total,
            purchase_date: purchaseDate,
            received_date: String(receivedValue).trim()
              ? excelDate(receivedValue)
              : "",
            import_batch: f.name,
          },
          fingerprint = rowFingerprint(
            body.bill_no,
            body.supplier_code || body.supplier,
            body.item_code || body.item,
            body.quantity,
            body.total,
            body.purchase_date,
          );
        if (seen.has(fingerprint)) {
          duplicates += 1;
          continue;
        }
        seen.add(fingerprint);
        inserts.push(body);
      }
      await bulkImport("purchases", inserts, [], "Purchase report");
      notify(
        `${inserts.length} purchase rows imported${duplicates ? ` · ${duplicates} duplicates skipped` : ""}${incomplete ? ` · ${incomplete} incomplete/summary rows skipped` : ""}`,
      );
      data.reload();
    } catch (e) {
      sendImportError("Purchase report");
      notify(e instanceof Error ? e.message : "Invalid Excel file");
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };
  const groupedBills = new Map<
    string,
    Record<string, unknown> & { items: DataRow[] }
  >();
  for (const row of data.rows) {
    const product = findProduct(products.rows, row.item_code, row.item),
      supplierMaster = suppliers.rows.find(
        (supplier) =>
          (row.supplier_code &&
            normalizedValue(supplier.code) ===
              normalizedValue(row.supplier_code)) ||
          (!row.supplier_code &&
            normalizedValue(supplier.name) === normalizedValue(row.supplier)),
      ),
      enrichedRow = {
        ...row,
        item: row.item || product?.name || `Product ${String(row.item_code)}`,
        item_code: row.item_code || product?.sku || "",
        rate: Number(row.rate || product?.purchase_rate || 0),
        supplier: row.supplier || supplierMaster?.name || row.supplier_code,
      },
      billNo = String(row.bill_no || `Purchase-${row.id}`),
      supplierKey = String(row.supplier_code || row.supplier || ""),
      groupKey = `${supplierKey}::${billNo}`,
      current = groupedBills.get(groupKey) || {
        bill_no: billNo,
        purchase_date: row.purchase_date,
        received_date: row.received_date,
        supplier: enrichedRow.supplier,
        supplier_code: row.supplier_code,
        total: 0,
        quantity: 0,
        items: [],
      };
    current.total = Number(current.total || 0) + Number(row.total || 0);
    current.quantity =
      Number(current.quantity || 0) + Number(row.quantity || 0);
    current.items.push(enrichedRow);
    groupedBills.set(groupKey, current);
  }
  const purchaseBills = [...groupedBills.values()]
    .filter((bill) => includesQuery(bill, query))
    .sort((a, b) =>
      sort === "newest" || sort === "oldest"
        ? String(a.purchase_date).localeCompare(String(b.purchase_date)) *
          (sort === "oldest" ? 1 : -1)
        : billCompare(
            a.bill_no,
            b.bill_no,
            sort === "billAsc" ? "asc" : "desc",
          ),
    );
  const removeBill = async (
    bill: Record<string, unknown> & { items: DataRow[] },
  ) => {
    if (!confirm(`Delete purchase bill ${String(bill.bill_no)}?`)) return;
    for (const item of bill.items)
      await request("purchases", "DELETE", undefined, item.id);
    notify("Purchase bill deleted");
    data.reload();
  };
  return (
    <div className="panel-pad">
      <div className="import-box">
        <span>⇩</span>
        <div>
          <h3>Import purchase details from Excel</h3>
          <p>
            Format: Supplier Code, Supplier Name, Bill No., Bill Date, Rec.Date,
            Code, Item, Qty and Product Value.
          </p>
          <small>
            {busy ? "Importing..." : file || "Upload .xlsx, .xls or .csv"}
          </small>
        </div>
        <label className="primary upload">
          Choose Excel
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={upload}
            disabled={busy}
          />
        </label>
        <button
          className="secondary"
          onClick={() =>
            exportExcel(
              data.rows.map((x) => ({
                "Supplier Code": x.supplier_code,
                "Supplier Name": x.supplier,
                "Bill No.": x.bill_no,
                "Bill Date": formatDate(x.purchase_date),
                "Rec.Date": formatDate(x.received_date),
                Code: x.item_code,
                Item: x.item,
                Qty: x.quantity,
                "Product Value": x.total,
              })),
              "purchases",
            )
          }
        >
          Export Excel
        </button>
      </div>
      <div className="smart-import-note">
        ✓ Supplier and product details are matched by code. Product Value ÷ Qty
        becomes the purchase rate and automatically updates missing Stock rates.
        Re-imported rows are skipped.
      </div>
      <div className="purchase-kpis">
        <article>
          <span>Total purchases</span>
          <b>{money(data.rows.reduce((s, x) => s + Number(x.total), 0))}</b>
        </article>
        <article>
          <span>Input GST</span>
          <b>
            {money(data.rows.reduce((s, x) => s + Number(x.gst_amount), 0))}
          </b>
        </article>
        <article>
          <span>Suppliers</span>
          <b>{new Set(data.rows.map((x) => x.supplier)).size}</b>
        </article>
        <article>
          <span>Purchase rows</span>
          <b>{data.rows.length}</b>
        </article>
      </div>
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search bill number, supplier, item or product code"
      >
        <label>
          Sort purchases
          <select
            value={sort}
            onChange={(event) =>
              setSort(
                event.target.value as
                  | "billAsc"
                  | "billDesc"
                  | "newest"
                  | "oldest",
              )
            }
          >
            <option value="billAsc">Bill No. A–Z</option>
            <option value="billDesc">Bill No. Z–A</option>
            <option value="newest">Date: new to old</option>
            <option value="oldest">Date: old to new</option>
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={purchaseBills.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (!confirm("Clear all purchase data?")) return;
          await clearResource("purchases");
          notify("Purchase data cleared");
          data.reload();
        }}
      />
      {purchaseBills.length ? (
        <div className="excel-preview">
          <table>
            <thead>
              <tr>
                <th>Bill</th>
                <th>Date</th>
                <th>Supplier</th>
                <th>Products</th>
                <th>Total Qty</th>
                <th>Total</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {purchaseBills.slice(0, limit).map((x) => (
                <tr key={`${String(x.supplier_code)}-${String(x.bill_no)}`}>
                  <td>{String(x.bill_no)}</td>
                  <td>{formatDate(x.purchase_date)}</td>
                  <td>{String(x.supplier)}</td>
                  <td>{x.items.length}</td>
                  <td>{String(x.quantity)}</td>
                  <td>{money(Number(x.total))}</td>
                  <td className="row-actions">
                    <button onClick={() => setViewingBill(x)}>View</button>
                    <button className="danger" onClick={() => removeBill(x)}>
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <Empty text="No purchase data. Import your Excel file to begin." />
      )}
      {viewingBill && (
        <BillDetails
          title={`Purchase Bill ${String(viewingBill.bill_no)}`}
          party={String(viewingBill.supplier || "Unknown supplier")}
          date={String(viewingBill.purchase_date || "")}
          items={viewingBill.items}
          close={() => setViewingBill(null)}
        />
      )}
    </div>
  );
}
function AccountsPanel({
  notify,
  mode,
}: {
  notify: (s: string) => void;
  mode: "cash" | "bank";
}) {
  const accounts = useResource("accounts"),
    tx = useResource("transactions"),
    [selected, setSelected] = useState<number | null>(null),
    [accountForm, setAccountForm] = useState(false),
    [txForm, setTxForm] = useState(false),
    [busy, setBusy] = useState(false),
    [editingAccount, setEditingAccount] = useState<DataRow | null>(null),
    [limit, setLimit] = useState(25),
    [query, setQuery] = useState(""),
    [directionFilter, setDirectionFilter] = useState("All");
  const visibleAccounts = accounts.rows.filter(
    (x) => String(x.type).toLowerCase() === mode,
  );
  useEffect(() => {
    if (!visibleAccounts.some((x) => x.id === selected))
      setSelected(visibleAccounts[0]?.id || null);
  }, [accounts.rows, mode, selected]);
  const balance = (id: number) =>
    Number(accounts.rows.find((x) => x.id === id)?.opening_balance || 0) +
    tx.rows
      .filter((x) => Number(x.account_id) === id)
      .reduce(
        (s, x) =>
          s + (x.direction === "in" ? Number(x.amount) : -Number(x.amount)),
        0,
      );
  const saveAccount = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    const body = {
      name: String(f.get("name")),
      type: mode === "cash" ? "Cash" : "Bank",
      bank_name: String(f.get("bank_name")),
      account_last4: String(f.get("account_last4")),
      opening_balance: Number(f.get("opening_balance")),
      active: 1,
    };
    await request(
      "accounts",
      editingAccount ? "PATCH" : "POST",
      body,
      editingAccount?.id,
    );
    notify("Account saved");
    setAccountForm(false);
    setEditingAccount(null);
    accounts.reload();
  };
  const saveTx = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    await request("transactions", "POST", {
      account_id: selected,
      direction: String(f.get("direction")),
      amount: Number(f.get("amount")),
      particulars: String(f.get("particulars")),
      reference: String(f.get("reference")),
      transaction_date: String(f.get("transaction_date")),
    });
    notify("Transaction saved");
    setTxForm(false);
    tx.reload();
  };
  const remove = async (id: number) => {
    await request("transactions", "DELETE", undefined, id);
    tx.reload();
  };
  const removeAccount = async () => {
    if (!selected || !confirm("Delete this account and all its transactions?"))
      return;
    for (const row of tx.rows.filter((x) => Number(x.account_id) === selected))
      await request("transactions", "DELETE", undefined, row.id);
    await request("accounts", "DELETE", undefined, selected);
    notify("Account deleted");
    setSelected(null);
    await Promise.all([accounts.reload(), tx.reload()]);
  };
  const importFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const rows = await readExcel(file);
      const known = [...visibleAccounts];
      const progressLabel = mode === "cash" ? "Cash account" : "Bank account";
      const transactions: Record<string, unknown>[] = [],
        seen = new Set(
          tx.rows.map((row) =>
            rowFingerprint(
              row.account_id,
              excelDate(row.transaction_date),
              row.direction,
              row.amount,
              row.particulars,
              row.reference,
            ),
          ),
        );
      let duplicates = 0,
        incomplete = 0;
      for (const row of rows) {
        const name = String(
          cell(row, "account", "account name", "bank account") ||
            (mode === "cash" ? "Cash Account" : "Bank Account"),
        );
        let account = known.find(
          (x) => String(x.name).toLowerCase() === name.toLowerCase(),
        );
        if (!account) {
          const created = await request("accounts", "POST", {
            name,
            type: mode === "cash" ? "Cash" : "Bank",
            bank_name: String(cell(row, "bank", "bank name")),
            account_last4: String(
              cell(row, "last 4", "last4", "account last4"),
            ),
            opening_balance: excelNumber(cell(row, "opening balance")),
            active: 1,
          });
          account = created.row;
          known.push(account as DataRow);
        }
        const debit = excelNumber(cell(row, "dr")),
          credit = excelNumber(cell(row, "cr")),
          direction = credit > 0 ? "out" : "in",
          body = {
            account_id: account.id,
            direction,
            amount: debit || credit,
            particulars: String(cell(row, "narration")),
            reference: String(cell(row, "#")),
            transaction_date: excelDate(cell(row, "date")),
          },
          fingerprint = rowFingerprint(
            body.account_id,
            body.transaction_date,
            body.direction,
            body.amount,
            body.particulars,
            body.reference,
          );
        if (!body.amount) {
          incomplete += 1;
          continue;
        }
        if (seen.has(fingerprint)) {
          duplicates += 1;
          continue;
        }
        seen.add(fingerprint);
        transactions.push(body);
      }
      await bulkImport("transactions", transactions, [], progressLabel);
      notify(
        `${transactions.length} ${mode} transactions imported${duplicates ? ` · ${duplicates} duplicates skipped` : ""}${incomplete ? ` · ${incomplete} incomplete rows skipped` : ""}`,
      );
      await Promise.all([accounts.reload(), tx.reload()]);
    } catch (err) {
      sendImportError(mode === "cash" ? "Cash account" : "Bank account");
      notify(err instanceof Error ? err.message : "Excel import failed");
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };
  const exportRows = () =>
    exportExcel(
      tx.rows
        .filter((x) =>
          visibleAccounts.some((a) => a.id === Number(x.account_id)),
        )
        .map((x, index) => ({
          "#": x.reference || index + 1,
          Date: formatDate(x.transaction_date),
          Account: visibleAccounts.find((a) => a.id === Number(x.account_id))
            ?.name,
          Narration: x.particulars,
          Dr: x.direction === "in" ? x.amount : 0,
          Cr: x.direction === "out" ? x.amount : 0,
          Balance: balance(Number(x.account_id)),
        })),
      mode === "cash" ? "cash-account" : "bank-accounts",
    );
  const filtered = tx.rows.filter(
    (x) =>
      Number(x.account_id) === selected &&
      (directionFilter === "All" || x.direction === directionFilter) &&
      includesQuery(x, query),
  );
  return (
    <div className="panel-pad">
      <div className="account-cards">
        {visibleAccounts.map((x) => (
          <button
            key={x.id}
            className={selected === x.id ? "active-account" : ""}
            onClick={() => setSelected(x.id)}
          >
            <span>{String(x.name)}</span>
            <b>{money(balance(x.id))}</b>
            <small>
              {String(x.type)}{" "}
              {x.account_last4 ? `· ${String(x.account_last4)}` : ""}
            </small>
          </button>
        ))}
        <button
          onClick={() => {
            setEditingAccount(null);
            setAccountForm(true);
          }}
        >
          <span>＋ Add {mode === "cash" ? "Cash" : "Bank"} Account</span>
          <b>{mode === "cash" ? "Cash ledger" : "Bank ledger"}</b>
          <small>Editable details</small>
        </button>
      </div>
      {accountForm && (
        <form className="invoice-builder" onSubmit={saveAccount}>
          <div className="form-grid">
            <label>
              Account Name
              <input
                name="name"
                defaultValue={String(editingAccount?.name || "")}
                required
              />
            </label>
            {mode === "bank" && (
              <label>
                Bank Name
                <input
                  name="bank_name"
                  defaultValue={String(editingAccount?.bank_name || "")}
                />
              </label>
            )}
            {mode === "bank" && (
              <label>
                Last 4 digits
                <input
                  name="account_last4"
                  maxLength={4}
                  defaultValue={String(editingAccount?.account_last4 || "")}
                />
              </label>
            )}
            <label>
              Opening Balance
              <input
                name="opening_balance"
                type="number"
                step=".01"
                defaultValue={Number(editingAccount?.opening_balance || 0)}
              />
            </label>
          </div>
          <div className="form-actions">
            <button className="primary">
              {editingAccount ? "Update" : "Save"} Account
            </button>
            <button
              type="button"
              onClick={() => {
                setAccountForm(false);
                setEditingAccount(null);
              }}
            >
              Cancel
            </button>
          </div>
        </form>
      )}
      <div className="account-actions">
        <b>{mode === "cash" ? "Cash" : "Bank"} Transactions</b>
        <div>
          <label className="secondary upload">
            {busy ? "Importing..." : "Import Excel"}
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={importFile}
              disabled={busy}
            />
          </label>
          <button onClick={exportRows}>Export Excel</button>
          <button
            disabled={!selected}
            onClick={() => {
              const row =
                visibleAccounts.find((x) => x.id === selected) || null;
              setEditingAccount(row);
              setAccountForm(true);
            }}
          >
            Edit Account
          </button>
          <button
            className="danger"
            disabled={!selected}
            onClick={removeAccount}
          >
            Delete Account
          </button>
          <button disabled={!selected} onClick={() => setTxForm(true)}>
            + Money In / Out
          </button>
        </div>
      </div>
      {txForm && (
        <form className="invoice-builder" onSubmit={saveTx}>
          <div className="form-grid">
            <label>
              Direction
              <select name="direction">
                <option value="in">Money In</option>
                <option value="out">Money Out</option>
              </select>
            </label>
            <label>
              Amount
              <input name="amount" type="number" step=".01" required />
            </label>
            <label>
              Particulars
              <input name="particulars" required />
            </label>
            <label>
              Reference
              <input name="reference" />
            </label>
            <label>
              Date
              <input
                name="transaction_date"
                type="date"
                defaultValue={new Date().toISOString().slice(0, 10)}
                required
              />
            </label>
          </div>
          <div className="form-actions">
            <button className="primary">Save Transaction</button>
          </div>
        </form>
      )}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder={`Search ${mode} narration, reference, date or amount`}
      >
        <label>
          Movement
          <select
            value={directionFilter}
            onChange={(event) => setDirectionFilter(event.target.value)}
          >
            <option value="All">All movements</option>
            <option value="in">Money In</option>
            <option value="out">Money Out</option>
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={filtered.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (!confirm(`Clear all ${mode} accounts and transactions?`)) return;
          await clearResource("transactions", `&type=${mode}`);
          await clearResource("accounts", `&type=${mode}`);
          notify(`${mode === "cash" ? "Cash" : "Bank"} data cleared`);
          await Promise.all([accounts.reload(), tx.reload()]);
        }}
      />
      {filtered.length ? (
        <table>
          <thead>
            <tr>
              <th>Date</th>
              <th>Particulars</th>
              <th>Money In</th>
              <th>Money Out</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {filtered.slice(0, limit).map((x) => (
              <tr key={x.id}>
                <td>{formatDate(x.transaction_date)}</td>
                <td>
                  <b>{String(x.particulars)}</b>
                  <small>{String(x.reference || "")}</small>
                </td>
                <td className="positive">
                  {x.direction === "in" ? money(Number(x.amount)) : "—"}
                </td>
                <td className="danger">
                  {x.direction === "out" ? money(Number(x.amount)) : "—"}
                </td>
                <td>
                  <button className="danger" onClick={() => remove(x.id)}>
                    Delete
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      ) : (
        <Empty text={`Add a ${mode} account or import Excel data to begin.`} />
      )}
    </div>
  );
}
function DataErrorsPanel({ notify }: { notify: (s: string) => void }) {
  const [audit, setAudit] = useState<{
      errors: Record<string, unknown>[];
      checked: Record<string, number>;
      generatedAt?: string;
    }>({ errors: [], checked: {} }),
    [loading, setLoading] = useState(true),
    [filter, setFilter] = useState("All"),
    [query, setQuery] = useState("");
  const load = async () => {
    setLoading(true);
    try {
      const response = await fetch("/api/audit"),
        data = await response.json();
      if (!response.ok) throw new Error(data.error || "Audit failed");
      setAudit(data);
      notify("Data correlation check completed");
    } catch (error) {
      notify(error instanceof Error ? error.message : "Audit failed");
    } finally {
      setLoading(false);
    }
  };
  useEffect(() => {
    load();
  }, []);
  const errors = audit.errors.filter(
    (x) =>
      (filter === "All" || x.severity === filter.toLowerCase()) &&
      includesQuery(x, query),
  );
  return (
    <div className="panel-pad">
      <div className="audit-summary">
        <article>
          <span>Errors</span>
          <b className="danger">
            {audit.errors.filter((x) => x.severity === "error").length}
          </b>
        </article>
        <article>
          <span>Warnings</span>
          <b>{audit.errors.filter((x) => x.severity === "warning").length}</b>
        </article>
        <article>
          <span>Records checked</span>
          <b>
            {Object.values(audit.checked)
              .reduce((a, b) => a + b, 0)
              .toLocaleString("en-IN")}
          </b>
        </article>
        <article>
          <span>Last checked</span>
          <b>
            {audit.generatedAt
              ? new Date(audit.generatedAt).toLocaleTimeString("en-IN")
              : "—"}
          </b>
        </article>
      </div>
      <div className="panel-tools">
        <div className="filter-tabs">
          {["All", "Error", "Warning"].map((x) => (
            <button
              key={x}
              className={filter === x ? "chosen" : ""}
              onClick={() => setFilter(x)}
            >
              {x}
            </button>
          ))}
        </div>
        <button className="primary" onClick={load}>
          {loading ? "Checking..." : "↻ Check Again"}
        </button>
      </div>
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search module, record or mismatch message"
      />
      {errors.length ? (
        <div className="error-list">
          {errors.map((x) => (
            <article key={String(x.id)} className={String(x.severity)}>
              <span>{x.severity === "error" ? "!" : "⚠"}</span>
              <div>
                <b>
                  {String(x.module)} · {String(x.record)}
                </b>
                <p>{String(x.message)}</p>
              </div>
              <small>{String(x.severity)}</small>
            </article>
          ))}
        </div>
      ) : (
        <Empty
          text={
            loading
              ? "Checking relationships across all data..."
              : "No mismatches found in the selected category."
          }
        />
      )}
    </div>
  );
}
function UsersPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("users"),
    [open, setOpen] = useState(false),
    [query, setQuery] = useState(""),
    [statusFilter, setStatusFilter] = useState("All");
  const visibleUsers = data.rows.filter(
    (user) =>
      (statusFilter === "All" ||
        (statusFilter === "Active" ? user.enabled : !user.enabled)) &&
      includesQuery(user, query),
  );
  const save = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    await request("users", "POST", {
      name: String(f.get("name")),
      email: String(f.get("email")),
      role: String(f.get("role")),
      enabled: 1,
    });
    notify("User added");
    setOpen(false);
    data.reload();
  };
  const toggle = async (u: DataRow) => {
    await request("users", "PATCH", { id: u.id, enabled: u.enabled ? 0 : 1 });
    notify(`${String(u.name)} ${u.enabled ? "disabled" : "enabled"}`);
    data.reload();
  };
  return (
    <div className="panel-pad">
      <div className="admin-note">
        <span>🛡</span>
        <div>
          <b>Admin authorization</b>
          <p>Disabled users are blocked by the backend API.</p>
        </div>
        <button className="primary" onClick={() => setOpen(!open)}>
          + Add User
        </button>
      </div>
      {open && (
        <form className="invoice-builder" onSubmit={save}>
          <div className="form-grid">
            <label>
              Name
              <input name="name" required />
            </label>
            <label>
              Email
              <input name="email" type="email" required />
            </label>
            <label>
              Role
              <select name="role">
                <option>Billing Staff</option>
                <option>Inventory Manager</option>
                <option>Viewer</option>
                <option>Admin</option>
              </select>
            </label>
          </div>
          <div className="form-actions">
            <button className="primary">Save User</button>
          </div>
        </form>
      )}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search user name, email or role"
      >
        <label>
          Status
          <select
            value={statusFilter}
            onChange={(event) => setStatusFilter(event.target.value)}
          >
            <option>All</option>
            <option>Active</option>
            <option>Disabled</option>
          </select>
        </label>
      </ModuleFilters>
      <div className="user-list">
        {visibleUsers.map((u) => (
          <article key={u.id}>
            <span className="user-avatar">
              {String(u.name)
                .split(" ")
                .map((x) => x[0])
                .slice(0, 2)
                .join("")}
            </span>
            <div>
              <b>{String(u.name)}</b>
              <small>{String(u.email)}</small>
            </div>
            <span className="role">{String(u.role)}</span>
            <span className={u.enabled ? "state enabled" : "state disabled"}>
              {u.enabled ? "Active" : "Disabled"}
            </span>
            <button
              className={u.role === "Admin" ? "disable locked" : "disable"}
              disabled={u.role === "Admin"}
              onClick={() => toggle(u)}
            >
              {u.role === "Admin" ? "Admin" : u.enabled ? "Disable" : "Enable"}
            </button>
          </article>
        ))}
      </div>
      {!data.rows.length && (
        <Empty text="The first signed-in user will be created as Admin." />
      )}
    </div>
  );
}
function InventoryPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("products"),
    purchases = useResource("purchases"),
    [edit, setEdit] = useState<DataRow | null>(null),
    [open, setOpen] = useState(false),
    [busy, setBusy] = useState(false),
    [limit, setLimit] = useState(25),
    [query, setQuery] = useState(""),
    [stockFilter, setStockFilter] = useState("All"),
    [sortOrder, setSortOrder] = useState<"newest" | "oldest">("newest");
  const latestPurchaseRates = buildPurchaseRateMap(purchases.rows),
    purchaseRateFor = (product: DataRow) =>
      effectivePurchaseRate(product, latestPurchaseRates);
  const visibleProducts = data.rows
    .filter((product) => {
      const stock = Number(product.stock),
        reorder = Number(product.reorder_level),
        matchesStock =
          stockFilter === "All" ||
          (stockFilter === "Low" && stock > 0 && stock <= reorder) ||
          (stockFilter === "Out" && stock <= 0) ||
          (stockFilter === "Available" && stock > reorder);
      return matchesStock && includesQuery(product, query);
    })
    .sort((a, b) => {
      const dateOrder = String(a.created_at || "").localeCompare(
        String(b.created_at || ""),
      );
      const order = dateOrder || a.id - b.id;
      return sortOrder === "oldest" ? order : -order;
    });
  const totalStockValue = data.rows.reduce(
    (sum, product) =>
      sum + Number(product.stock || 0) * purchaseRateFor(product),
    0,
  );
  const save = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget),
      body = {
        name: String(f.get("name")),
        sku: String(f.get("sku")),
        hsn_code: String(f.get("hsn_code")),
        category: String(f.get("category")),
        stock: Number(f.get("stock")),
        reorder_level: Number(f.get("reorder_level")),
        purchase_rate: Number(f.get("purchase_rate")),
        price: Number(f.get("price")),
        gst_rate: Number(f.get("gst_rate")),
        created_at: String(edit?.created_at || new Date().toISOString()),
      };
    await request(
      "products",
      edit ? "PATCH" : "POST",
      edit ? { ...body, id: edit.id } : body,
    );
    notify(edit ? "Product updated" : "Product added");
    setOpen(false);
    setEdit(null);
    data.reload();
  };
  const remove = async (id: number) => {
    if (confirm("Delete product?")) {
      await request("products", "DELETE", undefined, id);
      data.reload();
    }
  };
  const importFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const allRows = await readExcel(file),
        validRows = allRows.filter((row) =>
          String(cell(row, "item code")).trim(),
        ),
        uniqueRows = new Map<string, Record<string, unknown>>();
      for (const row of validRows)
        uniqueRows.set(
          String(cell(row, "item code")).trim().toLowerCase(),
          row,
        );
      const rows = [...uniqueRows.values()],
        skipped = allRows.length - validRows.length,
        fileDuplicates = validRows.length - rows.length;
      const inserts: Record<string, unknown>[] = [],
        updates: Record<string, unknown>[] = [];
      let unchanged = 0;
      for (const row of rows) {
        const sku = String(cell(row, "item code")).trim(),
          enteredName = String(cell(row, "item name")).trim(),
          enteredHsn = String(cell(row, "hsn code")).trim(),
          enteredGst = cell(row, "gst %"),
          enteredMrp = cell(row, "mrp"),
          enteredStock = cell(row, "stock");
        const existing = data.rows.find(
          (x) => String(x.sku).trim().toLowerCase() === sku.toLowerCase(),
        );
        const body = {
          name: enteredName || String(existing?.name || `Product ${sku}`),
          sku,
          hsn_code: enteredHsn || String(existing?.hsn_code || ""),
          category: String(existing?.category || "General"),
          gst_rate: String(enteredGst).trim()
            ? excelNumber(enteredGst)
            : Number(existing?.gst_rate || 0),
          purchase_rate:
            excelNumber(
              cell(row, "purchase rate", "purchase price", "cost price"),
            ) ||
            purchaseRateFor(
              existing ||
                ({
                  id: 0,
                  sku,
                  name: enteredName || `Product ${sku}`,
                } as DataRow),
            ),
          price: String(enteredMrp).trim()
            ? excelNumber(enteredMrp)
            : Number(existing?.price || 0),
          stock: String(enteredStock).trim()
            ? excelNumber(enteredStock)
            : Number(existing?.stock || 0),
          reorder_level: Number(existing?.reorder_level || 5),
          created_at: String(existing?.created_at || new Date().toISOString()),
        };
        if (
          existing &&
          sameFields(existing, body, [
            "name",
            "sku",
            "hsn_code",
            "category",
            "gst_rate",
            "purchase_rate",
            "price",
            "stock",
            "reorder_level",
          ])
        )
          unchanged += 1;
        else if (existing) updates.push({ ...body, id: existing.id });
        else inserts.push(body);
      }
      await bulkImport("products", inserts, updates, "Stock master");
      notify(
        `${inserts.length} new and ${updates.length} changed stock items imported${fileDuplicates + unchanged ? ` · ${fileDuplicates + unchanged} duplicates skipped` : ""}${skipped ? ` · ${skipped} rows without product code skipped` : ""}`,
      );
      data.reload();
    } catch (err) {
      sendImportError("Stock master");
      notify(err instanceof Error ? err.message : "Stock import failed");
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };
  return (
    <div className="panel-pad">
      <div className="mini-stats">
        <b>{data.rows.length} Products</b>
        <b>
          {data.rows.reduce((s, x) => s + Number(x.stock), 0)} Units Available
        </b>
        <b>{money(totalStockValue)} Stock Value</b>
        <b className="danger">
          {
            data.rows.filter((x) => Number(x.stock) <= Number(x.reorder_level))
              .length
          }{" "}
          Stock Alerts
        </b>
      </div>
      <div className="excel-actions return-add">
        <button
          className="primary"
          onClick={() => {
            setEdit(null);
            setOpen(true);
          }}
        >
          + Add Product
        </button>
        <label className="secondary upload">
          {busy ? "Importing..." : "Import Stock Excel"}
          <input type="file" accept=".xlsx,.xls,.csv" onChange={importFile} />
        </label>
        <button
          className="secondary"
          onClick={() =>
            exportExcel(
              data.rows.map((x) => ({
                "Item Code": x.sku,
                "Item Name": x.name,
                "HSN Code": x.hsn_code,
                "GST %": x.gst_rate,
                "Purchase Rate": purchaseRateFor(x),
                "Selling Rate": x.price,
                Stock: x.stock,
                "Stock Value": Number(x.stock) * purchaseRateFor(x),
                "Added Date": formatDate(x.created_at),
              })),
              "stock",
            )
          }
        >
          Export Excel
        </button>
      </div>
      <div className="smart-import-note">
        ✓ Product code is the master key. Missing purchase rates are fetched
        from the latest Purchase report; missing selling rates are fetched from
        Sales. Files may be imported in any order and unchanged rows are
        skipped.
      </div>
      {open && (
        <form className="invoice-builder" onSubmit={save}>
          <div className="form-grid">
            {[
              ["name", "Product name"],
              ["sku", "SKU"],
              ["hsn_code", "HSN Code"],
              ["category", "Category"],
              ["stock", "Stock"],
              ["reorder_level", "Reorder level"],
              ["purchase_rate", "Purchase rate"],
              ["price", "Selling rate"],
              ["gst_rate", "GST rate"],
            ].map(([n, l]) => (
              <label key={n}>
                {l} *
                <input
                  name={n}
                  type={
                    [
                      "stock",
                      "reorder_level",
                      "purchase_rate",
                      "price",
                      "gst_rate",
                    ].includes(n)
                      ? "number"
                      : "text"
                  }
                  step=".01"
                  defaultValue={String(
                    n === "purchase_rate" && edit
                      ? purchaseRateFor(edit)
                      : (edit?.[n] ?? (n === "gst_rate" ? 18 : "")),
                  )}
                  required
                />
              </label>
            ))}
          </div>
          <div className="form-actions">
            <button
              type="button"
              className="secondary"
              onClick={() => setOpen(false)}
            >
              Cancel
            </button>
            <button className="primary">Save Product</button>
          </div>
        </form>
      )}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search item name, SKU, HSN or category"
      >
        <label>
          Stock status
          <select
            value={stockFilter}
            onChange={(event) => setStockFilter(event.target.value)}
          >
            <option value="All">All items</option>
            <option value="Low">Low stock</option>
            <option value="Out">Out of stock</option>
            <option value="Available">Available</option>
          </select>
        </label>
        <label>
          Added order
          <select
            value={sortOrder}
            onChange={(event) =>
              setSortOrder(event.target.value as "newest" | "oldest")
            }
          >
            <option value="newest">New to old</option>
            <option value="oldest">Old to new</option>
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={visibleProducts.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (!confirm("Clear all stock data?")) return;
          await clearResource("products");
          notify("Stock data cleared");
          data.reload();
        }}
      />
      <div className="inventory-list">
        {visibleProducts.slice(0, limit).map((x) => (
          <article key={x.id}>
            <div>
              <b>{String(x.name)}</b>
              <small>
                {String(x.category)} · SKU {String(x.sku)} · Reorder at{" "}
                {String(x.reorder_level)} · Added {formatDate(x.created_at)}
              </small>
            </div>
            <div className="stock-rates">
              <span>
                Purchase <b>{money(purchaseRateFor(x))}</b>
              </span>
              <span>
                Selling <b>{money(Number(x.price || 0))}</b>
              </span>
              <span>
                Stock value{" "}
                <b>{money(Number(x.stock || 0) * purchaseRateFor(x))}</b>
              </span>
            </div>
            <strong
              className={
                Number(x.stock) <= Number(x.reorder_level) ? "danger" : ""
              }
            >
              {String(x.stock)}
            </strong>
            <div className="row-actions">
              <button
                onClick={() => {
                  setEdit(x);
                  setOpen(true);
                }}
              >
                Edit
              </button>
              <button className="danger" onClick={() => remove(x.id)}>
                Delete
              </button>
            </div>
          </article>
        ))}
      </div>
      {!visibleProducts.length && (
        <Empty text="No stock items match the current search and filter." />
      )}
    </div>
  );
}
function PaymentsPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("payments"),
    customers = useResource("customers"),
    invoices = useResource("invoices"),
    [open, setOpen] = useState(false),
    [query, setQuery] = useState(""),
    [methodFilter, setMethodFilter] = useState("All"),
    [limit, setLimit] = useState(25);
  const visiblePayments = data.rows.filter((payment) => {
    const customer = customers.rows.find(
      (item) => item.id === Number(payment.customer_id),
    );
    return (
      (methodFilter === "All" || payment.method === methodFilter) &&
      includesQuery({ ...payment, customer_name: customer?.name || "" }, query)
    );
  });
  const save = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    await request("payments", "POST", {
      customer_id: Number(f.get("customer_id")),
      invoice_id: f.get("invoice_id") ? Number(f.get("invoice_id")) : null,
      amount: Number(f.get("amount")),
      method: String(f.get("method")),
      reference: String(f.get("reference")),
      collected_at: String(f.get("collected_at")),
    });
    notify("Payment saved");
    setOpen(false);
    data.reload();
  };
  const remove = async (id: number) => {
    await request("payments", "DELETE", undefined, id);
    data.reload();
  };
  return (
    <div className="panel-pad">
      <div className="payment-summary">
        <article>
          <span>Total collected</span>
          <b>{money(data.rows.reduce((s, x) => s + Number(x.amount), 0))}</b>
        </article>
        <article>
          <span>Transactions</span>
          <b>{data.rows.length}</b>
        </article>
        <article>
          <span>Storage</span>
          <b>Live database</b>
        </article>
      </div>
      <button className="primary" onClick={() => setOpen(!open)}>
        ₹ Record Payment
      </button>
      {open && (
        <form className="invoice-builder" onSubmit={save}>
          <div className="form-grid">
            <label>
              Customer
              <select name="customer_id" required>
                <option value="">Select customer</option>
                {customers.rows.map((x) => (
                  <option key={x.id} value={x.id}>
                    {String(x.code || "")} {String(x.name)}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Invoice
              <select name="invoice_id">
                <option value="">General payment</option>
                {invoices.rows
                  .filter((x) => x.kind !== "quotation")
                  .map((x) => (
                    <option key={x.id} value={x.id}>
                      {String(x.invoice_no)} · {String(x.customer_name)}
                    </option>
                  ))}
              </select>
            </label>
            <label>
              Amount
              <input name="amount" type="number" step=".01" required />
            </label>
            <label>
              Method
              <select name="method">
                <option>Cash</option>
                <option>UPI</option>
                <option>Bank Transfer</option>
                <option>Cheque</option>
                <option>Card</option>
              </select>
            </label>
            <label>
              Reference
              <input name="reference" />
            </label>
            <label>
              Date
              <input
                name="collected_at"
                type="date"
                defaultValue={new Date().toISOString().slice(0, 10)}
                required
              />
            </label>
          </div>
          <div className="form-actions">
            <button className="primary">Save Payment</button>
          </div>
        </form>
      )}{" "}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search customer, reference, date or amount"
      >
        <label>
          Method
          <select
            value={methodFilter}
            onChange={(event) => setMethodFilter(event.target.value)}
          >
            <option>All</option>
            <option>Cash</option>
            <option>UPI</option>
            <option>Bank Transfer</option>
            <option>Cheque</option>
            <option>Card</option>
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={visiblePayments.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (!confirm("Clear all payment entries?")) return;
          await clearResource("payments");
          notify("Payment data cleared");
          data.reload();
        }}
      />
      {visiblePayments.length ? (
        <div className="excel-preview">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Customer</th>
                <th>Method</th>
                <th>Reference</th>
                <th>Amount</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {visiblePayments.slice(0, limit).map((x) => (
                <tr key={x.id}>
                  <td>{formatDate(x.collected_at)}</td>
                  <td>
                    {String(
                      customers.rows.find((c) => c.id === Number(x.customer_id))
                        ?.name || `#${String(x.customer_id)}`,
                    )}
                  </td>
                  <td>{String(x.method)}</td>
                  <td>{String(x.reference || "")}</td>
                  <td>{money(Number(x.amount))}</td>
                  <td>
                    <button className="danger" onClick={() => remove(x.id)}>
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <Empty text="No payment entries yet." />
      )}
    </div>
  );
}
function QuotationPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("invoices"),
    [open, setOpen] = useState(false),
    [editing, setEditing] = useState<DataRow | null>(null),
    [query, setQuery] = useState(""),
    [statusFilter, setStatusFilter] = useState("All"),
    [limit, setLimit] = useState(25);
  const quotes = data.rows
      .filter((x) => x.kind === "quotation")
      .filter(
        (quote) =>
          (statusFilter === "All" || quote.status === statusFilter) &&
          includesQuery(quote, query),
      ),
    quotationStatuses = [
      "All",
      ...new Set(
        data.rows
          .filter((x) => x.kind === "quotation")
          .map((x) => String(x.status || "Draft")),
      ),
    ];
  const save = async (body: Record<string, unknown>) => {
    await request(
      "invoices",
      editing ? "PATCH" : "POST",
      editing ? { ...body, id: editing.id } : body,
    );
    notify(editing ? "Quotation updated" : "Quotation saved");
    setOpen(false);
    setEditing(null);
    data.reload();
  };
  const remove = async (id: number) => {
    if (!confirm("Delete quotation?")) return;
    await request("invoices", "DELETE", undefined, id);
    notify("Quotation deleted");
    data.reload();
  };
  const convert = async (row: DataRow) => {
    await request("invoices", "PATCH", {
      ...row,
      id: row.id,
      kind: "invoice",
      invoice_no: String(row.invoice_no).replace(/^QT-/, "INV-"),
      status: "Pending",
    });
    notify("Quotation converted to sales invoice");
    data.reload();
  };
  return (
    <div className="panel-pad">
      <button
        className="primary"
        onClick={() => {
          setEditing(null);
          setOpen(!open);
        }}
      >
        {open ? "Close Form" : "+ Create Quotation"}
      </button>
      {open && (
        <InvoiceBuilder
          initial={editing}
          save={save}
          close={() => setOpen(false)}
          kind="quotation"
        />
      )}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search quotation, customer, product or amount"
      >
        <label>
          Status
          <select
            value={statusFilter}
            onChange={(event) => setStatusFilter(event.target.value)}
          >
            {quotationStatuses.map((status) => (
              <option key={status}>{status}</option>
            ))}
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={quotes.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (!confirm("Clear all quotations?")) return;
          await clearResource("invoices", "&kind=quotation");
          notify("Quotation data cleared");
          data.reload();
        }}
      />
      {quotes.length ? (
        <div className="excel-preview">
          <table>
            <thead>
              <tr>
                <th>Quotation</th>
                <th>Customer</th>
                <th>Date</th>
                <th>Amount</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {quotes.slice(0, limit).map((x) => (
                <tr key={x.id}>
                  <td>{String(x.invoice_no)}</td>
                  <td>{String(x.customer_name)}</td>
                  <td>{formatDate(x.created_at)}</td>
                  <td>{money(Number(x.amount))}</td>
                  <td>{String(x.status)}</td>
                  <td className="row-actions">
                    <button
                      onClick={() => {
                        setEditing(x);
                        setOpen(true);
                      }}
                    >
                      Edit
                    </button>
                    <button onClick={() => convert(x)}>
                      Convert to Invoice
                    </button>
                    <button className="danger" onClick={() => remove(x.id)}>
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <Empty text="No quotations yet." />
      )}
    </div>
  );
}
function CustomersPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("customers"),
    [open, setOpen] = useState(false),
    [edit, setEdit] = useState<DataRow | null>(null),
    [busy, setBusy] = useState(false),
    [limit, setLimit] = useState(25),
    [query, setQuery] = useState(""),
    [registrationFilter, setRegistrationFilter] = useState("All");
  const customerTypes = [
      "All",
      ...new Set(
        data.rows
          .map((row) => String(row.registration_type || ""))
          .filter(Boolean),
      ),
    ],
    visibleCustomers = data.rows.filter(
      (customer) =>
        (registrationFilter === "All" ||
          customer.registration_type === registrationFilter) &&
        includesQuery(customer, query),
    );
  const save = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget),
      body = {
        name: String(f.get("name")),
        owner_name: String(f.get("owner_name")),
        code: String(f.get("code")),
        gstin: String(f.get("gstin")),
        phone: String(f.get("phone")),
        registration_type: String(f.get("registration_type")),
        state: String(f.get("state")),
        email: String(f.get("email")),
        address: String(f.get("address")),
        balance: Number(f.get("balance")),
      };
    await request(
      "customers",
      edit ? "PATCH" : "POST",
      edit ? { ...body, id: edit.id } : body,
    );
    setOpen(false);
    data.reload();
  };
  const remove = async (id: number) => {
    await request("customers", "DELETE", undefined, id);
    data.reload();
  };
  const importFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const rows = await readExcel(file);
      const inserts: Record<string, unknown>[] = [],
        updates: Record<string, unknown>[] = [],
        fileFingerprints = new Set<string>();
      let unchanged = 0,
        incomplete = 0;
      for (const row of rows) {
        const code = String(cell(row, "code")).trim(),
          enteredName = String(cell(row, "shop name")).trim(),
          existing = data.rows.find(
            (x) =>
              (code && normalizedValue(x.code) === normalizedValue(code)) ||
              (!code &&
                enteredName &&
                normalizedValue(x.name) === normalizedValue(enteredName)),
          ),
          body = {
            name: enteredName || String(existing?.name || ""),
            owner_name:
              String(cell(row, "owner name")).trim() ||
              String(existing?.owner_name || ""),
            code: code || String(existing?.code || ""),
            phone:
              String(cell(row, "contact no")).trim() ||
              String(existing?.phone || ""),
            registration_type:
              String(cell(row, "reg/unreg")).trim() ||
              String(existing?.registration_type || ""),
            state:
              String(cell(row, "state")).trim() ||
              String(existing?.state || ""),
            email:
              String(cell(row, "email")).trim() ||
              String(existing?.email || ""),
            gstin:
              String(cell(row, "gstin")).trim() ||
              String(existing?.gstin || ""),
            address:
              String(cell(row, "address")).trim() ||
              String(existing?.address || ""),
            balance: Number(existing?.balance || 0),
          };
        if (!body.name) {
          incomplete += 1;
          continue;
        }
        const fingerprint = rowFingerprint(
          body.name,
          body.owner_name,
          body.code,
          body.phone,
          body.registration_type,
          body.state,
          body.email,
          body.gstin,
          body.address,
        );
        if (fileFingerprints.has(fingerprint)) {
          unchanged += 1;
          continue;
        }
        fileFingerprints.add(fingerprint);
        if (
          existing &&
          sameFields(existing, body, [
            "name",
            "owner_name",
            "code",
            "phone",
            "registration_type",
            "state",
            "email",
            "gstin",
            "address",
            "balance",
          ])
        )
          unchanged += 1;
        else if (existing) updates.push({ ...body, id: existing.id });
        else inserts.push(body);
      }
      await bulkImport("customers", inserts, updates, "Customer master");
      notify(
        `${inserts.length} new and ${updates.length} changed customers imported${unchanged ? ` · ${unchanged} duplicates skipped` : ""}${incomplete ? ` · ${incomplete} incomplete rows skipped` : ""}`,
      );
      data.reload();
    } catch (err) {
      sendImportError("Customer master");
      notify(err instanceof Error ? err.message : "Customer import failed");
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };
  return (
    <div className="panel-pad">
      <div className="excel-actions">
        <button
          className="primary"
          onClick={() => {
            setEdit(null);
            setOpen(true);
          }}
        >
          + Add Customer
        </button>
        <label className="secondary upload">
          {busy ? "Importing..." : "Import Customer Excel"}
          <input type="file" accept=".xlsx,.xls,.csv" onChange={importFile} />
        </label>
        <button
          className="secondary"
          onClick={() =>
            exportExcel(
              data.rows.map((x) => ({
                "Shop Name": x.name,
                "Owner Name": x.owner_name,
                Code: x.code,
                "Contact No": x.phone,
                "Reg/Unreg": x.registration_type,
                State: x.state,
                Email: x.email,
                GSTIN: x.gstin,
                Address: x.address,
              })),
              "customers",
            )
          }
        >
          Export Excel
        </button>
      </div>
      {open && (
        <form className="invoice-builder" onSubmit={save}>
          <div className="form-grid">
            <label>
              Name
              <input
                name="name"
                defaultValue={String(edit?.name || "")}
                required
              />
            </label>
            <label>
              GSTIN
              <input name="gstin" defaultValue={String(edit?.gstin || "")} />
            </label>
            {[
              ["owner_name", "Owner Name"],
              ["code", "Code"],
              ["registration_type", "Reg/Unreg"],
              ["state", "State"],
              ["email", "Email"],
              ["address", "Address"],
            ].map(([n, l]) => (
              <label key={n}>
                {l}
                <input name={n} defaultValue={String(edit?.[n] || "")} />
              </label>
            ))}
            <label>
              Phone
              <input name="phone" defaultValue={String(edit?.phone || "")} />
            </label>
            <label>
              Opening Balance
              <input
                name="balance"
                type="number"
                step=".01"
                defaultValue={String(edit?.balance || 0)}
              />
            </label>
          </div>
          <div className="form-actions">
            <button className="primary">Save Customer</button>
          </div>
        </form>
      )}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search shop, owner, code, contact, GSTIN or state"
      >
        <label>
          Registration
          <select
            value={registrationFilter}
            onChange={(event) => setRegistrationFilter(event.target.value)}
          >
            {customerTypes.map((type) => (
              <option key={type}>{type}</option>
            ))}
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={visibleCustomers.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (
            !confirm(
              "Clear all customer data? Existing invoice records will remain.",
            )
          )
            return;
          await clearResource("customers");
          notify("Customer data cleared");
          data.reload();
        }}
      />
      <div className="feature-grid">
        {visibleCustomers.slice(0, limit).map((x) => (
          <article key={x.id}>
            <span>♧</span>
            <h3>{String(x.name)}</h3>
            <p>
              {String(x.phone || "No phone")} · {String(x.gstin || "No GSTIN")}
            </p>
            <b>Balance: {money(Number(x.balance))}</b>
            <div className="row-actions">
              <button
                onClick={() => {
                  setEdit(x);
                  setOpen(true);
                }}
              >
                Edit
              </button>
              <button className="danger" onClick={() => remove(x.id)}>
                Delete
              </button>
            </div>
          </article>
        ))}
      </div>
      {!visibleCustomers.length && (
        <Empty text="No customers match the current search and filter." />
      )}
    </div>
  );
}
function SuppliersPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("suppliers"),
    [busy, setBusy] = useState(false),
    [limit, setLimit] = useState(25),
    [open, setOpen] = useState(false),
    [editing, setEditing] = useState<DataRow | null>(null),
    [query, setQuery] = useState(""),
    [stateFilter, setStateFilter] = useState("All");
  const supplierStates = [
      "All",
      ...new Set(
        data.rows
          .map((row) => String(row.state || ""))
          .filter(Boolean)
          .sort(),
      ),
    ],
    visibleSuppliers = data.rows.filter(
      (supplier) =>
        (stateFilter === "All" || supplier.state === stateFilter) &&
        includesQuery(supplier, query),
    );
  const save = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget),
      body = {
        name: String(f.get("name")),
        code: String(f.get("code")),
        phone: String(f.get("phone")),
        state: String(f.get("state")),
        email: String(f.get("email")),
        gstin: String(f.get("gstin")),
        address: String(f.get("address")),
      };
    await request("suppliers", editing ? "PATCH" : "POST", body, editing?.id);
    notify(editing ? "Supplier updated" : "Supplier added");
    setOpen(false);
    setEditing(null);
    data.reload();
  };
  const importFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const rows = await readExcel(file);
      const inserts: Record<string, unknown>[] = [],
        updates: Record<string, unknown>[] = [],
        fileFingerprints = new Set<string>();
      let unchanged = 0,
        incomplete = 0;
      for (const row of rows) {
        const code = String(cell(row, "code")).trim(),
          enteredName = String(cell(row, "name")).trim(),
          existing = data.rows.find(
            (x) =>
              (code && normalizedValue(x.code) === normalizedValue(code)) ||
              (!code &&
                enteredName &&
                normalizedValue(x.name) === normalizedValue(enteredName)),
          ),
          body = {
            name: enteredName || String(existing?.name || ""),
            code: code || String(existing?.code || ""),
            phone:
              String(cell(row, "contact no")).trim() ||
              String(existing?.phone || ""),
            state:
              String(cell(row, "state")).trim() ||
              String(existing?.state || ""),
            email:
              String(cell(row, "email")).trim() ||
              String(existing?.email || ""),
            gstin:
              String(cell(row, "gstin")).trim() ||
              String(existing?.gstin || ""),
            address:
              String(cell(row, "address")).trim() ||
              String(existing?.address || ""),
          };
        if (!body.name) {
          incomplete += 1;
          continue;
        }
        const fingerprint = rowFingerprint(
          body.name,
          body.code,
          body.phone,
          body.state,
          body.email,
          body.gstin,
          body.address,
        );
        if (fileFingerprints.has(fingerprint)) {
          unchanged += 1;
          continue;
        }
        fileFingerprints.add(fingerprint);
        if (
          existing &&
          sameFields(existing, body, [
            "name",
            "code",
            "phone",
            "state",
            "email",
            "gstin",
            "address",
          ])
        )
          unchanged += 1;
        else if (existing) updates.push({ ...body, id: existing.id });
        else inserts.push(body);
      }
      await bulkImport("suppliers", inserts, updates, "Supplier master");
      notify(
        `${inserts.length} new and ${updates.length} changed suppliers imported${unchanged ? ` · ${unchanged} duplicates skipped` : ""}${incomplete ? ` · ${incomplete} incomplete rows skipped` : ""}`,
      );
      data.reload();
    } catch (err) {
      sendImportError("Supplier master");
      notify(err instanceof Error ? err.message : "Supplier import failed");
    } finally {
      setBusy(false);
      e.target.value = "";
    }
  };
  return (
    <div className="panel-pad">
      <div className="excel-actions">
        <button
          className="primary"
          onClick={() => {
            setEditing(null);
            setOpen(!open);
          }}
        >
          {open ? "Close Form" : "+ Add Supplier"}
        </button>
        <label className="primary upload">
          {busy ? "Importing..." : "Import Supplier Excel"}
          <input type="file" accept=".xlsx,.xls,.csv" onChange={importFile} />
        </label>
        <button
          className="secondary"
          onClick={() =>
            exportExcel(
              data.rows.map((x) => ({
                Name: x.name,
                Code: x.code,
                "Contact No": x.phone,
                State: x.state,
                Email: x.email,
                GSTIN: x.gstin,
                Address: x.address,
              })),
              "suppliers",
            )
          }
        >
          Export Excel
        </button>
      </div>
      {open && (
        <form className="invoice-builder" onSubmit={save}>
          <div className="form-grid">
            {[
              ["name", "Supplier Name"],
              ["code", "Code"],
              ["phone", "Contact No"],
              ["state", "State"],
              ["email", "Email"],
              ["gstin", "GSTIN"],
              ["address", "Address"],
            ].map(([name, label]) => (
              <label key={name}>
                {label}
                <input
                  name={name}
                  defaultValue={String(editing?.[name] || "")}
                  required={name === "name"}
                />
              </label>
            ))}
          </div>
          <div className="form-actions">
            <button type="button" onClick={() => setOpen(false)}>
              Cancel
            </button>
            <button className="primary">Save Supplier</button>
          </div>
        </form>
      )}
      <ModuleFilters
        query={query}
        setQuery={setQuery}
        placeholder="Search supplier, code, contact, GSTIN or email"
      >
        <label>
          State
          <select
            value={stateFilter}
            onChange={(event) => setStateFilter(event.target.value)}
          >
            {supplierStates.map((state) => (
              <option key={state}>{state}</option>
            ))}
          </select>
        </label>
      </ModuleFilters>
      <DataControls
        total={visibleSuppliers.length}
        limit={limit}
        setLimit={setLimit}
        clear={async () => {
          if (!confirm("Clear all supplier data?")) return;
          await clearResource("suppliers");
          notify("Supplier data cleared");
          data.reload();
        }}
      />
      {visibleSuppliers.length ? (
        <div className="excel-preview">
          <table>
            <thead>
              <tr>
                <th>Code</th>
                <th>Name</th>
                <th>Contact</th>
                <th>State</th>
                <th>GSTIN</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {visibleSuppliers.slice(0, limit).map((x) => (
                <tr key={x.id}>
                  <td>{String(x.code || "")}</td>
                  <td>{String(x.name)}</td>
                  <td>{String(x.phone || "")}</td>
                  <td>{String(x.state || "")}</td>
                  <td>{String(x.gstin || "")}</td>
                  <td>
                    <button
                      onClick={() => {
                        setEditing(x);
                        setOpen(true);
                      }}
                    >
                      Edit
                    </button>{" "}
                    <button
                      className="danger"
                      onClick={async () => {
                        await request("suppliers", "DELETE", undefined, x.id);
                        data.reload();
                      }}
                    >
                      Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <Empty text="No suppliers yet. Import the supplier Excel format to begin." />
      )}
    </div>
  );
}
function ReportsPanel() {
  const [s, setS] = useState<Record<string, number>>({});
  const [dealerQuery, setDealerQuery] = useState("");
  const [productQuery, setProductQuery] = useState("");
  const [audit, setAudit] = useState<{
    bestSellers: Record<string, unknown>[];
    bestDealers: Record<string, unknown>[];
    stockValue: number;
  }>({ bestSellers: [], bestDealers: [], stockValue: 0 });
  const invoices = useResource("invoices"),
    purchasesData = useResource("purchases"),
    returnsData = useResource("returns"),
    transactions = useResource("transactions"),
    accounts = useResource("accounts"),
    products = useResource("products");
  useEffect(() => {
    fetch("/api/summary")
      .then((r) => r.json())
      .then(setS);
    fetch("/api/audit")
      .then((r) => r.json())
      .then((data) =>
        setAudit({
          bestSellers: data.bestSellers || [],
          bestDealers: data.bestDealers || [],
          stockValue: Number(data.stockValue || 0),
        }),
      )
      .catch(() => {});
  }, []);
  const max = Math.max(s.sales || 0, s.purchases || 0, 1),
    sales = Math.max(3, Math.round(((s.sales || 0) / max) * 90)),
    purchases = Math.max(3, Math.round(((s.purchases || 0) / max) * 90)),
    profit = (s.sales || 0) - (s.purchases || 0) - (s.returns || 0),
    netSales = (s.sales || 0) - (s.returns || 0),
    profitMargin = netSales ? (profit / netSales) * 100 : 0,
    reportPurchaseRates = buildPurchaseRateMap(purchasesData.rows),
    stockValue = products.rows.length
      ? products.rows.reduce(
          (sum, product) =>
            sum +
            Number(product.stock || 0) *
              effectivePurchaseRate(product, reportPurchaseRates),
          0,
        )
      : audit.stockValue,
    visibleDealers = audit.bestDealers.filter((dealer) =>
      includesQuery(dealer, dealerQuery),
    ),
    visibleBestSellers = audit.bestSellers.filter((product) =>
      includesQuery(product, productQuery),
    );
  const downloadReport = (name: string) => {
    let rows: Record<string, unknown>[] = [];
    if (name === "Sales Register")
      rows = invoices.rows
        .filter((x) => x.kind !== "quotation")
        .map((x) => ({ ...x, created_at: formatDate(x.created_at) }));
    else if (name === "Purchase Register")
      rows = purchasesData.rows.map((x) => ({
        ...x,
        purchase_date: formatDate(x.purchase_date),
        received_date: formatDate(x.received_date),
      }));
    else if (name === "Sales Return Register")
      rows = returnsData.rows.map((x) => ({
        ...x,
        created_at: formatDate(x.created_at),
      }));
    else if (name === "GST Summary")
      rows = [
        {
          Output_GST: invoices.rows.reduce((a, x) => a + Number(x.tax || 0), 0),
          Input_GST: purchasesData.rows.reduce(
            (a, x) => a + Number(x.gst_amount || 0),
            0,
          ),
        },
      ];
    else if (name === "Cash Book" || name === "Bank Book") {
      const type = name.startsWith("Cash") ? "cash" : "bank",
        ids = accounts.rows
          .filter((x) => String(x.type).toLowerCase() === type)
          .map((x) => x.id);
      rows = transactions.rows
        .filter((x) => ids.includes(Number(x.account_id)))
        .map((x) => ({
          ...x,
          transaction_date: formatDate(x.transaction_date),
        }));
    } else if (name === "Outstanding Report")
      rows = invoices.rows
        .filter(
          (x) =>
            x.kind !== "quotation" && String(x.status).toLowerCase() !== "paid",
        )
        .map((x) => ({ ...x, created_at: formatDate(x.created_at) }));
    else if (name === "Stock Valuation")
      rows = products.rows.map((x) => ({
        "Item Code": x.sku,
        "Item Name": x.name,
        Stock: x.stock,
        "Purchase Rate": effectivePurchaseRate(x, reportPurchaseRates),
        "Selling Rate": x.price,
        "Stock Value":
          Number(x.stock) * effectivePurchaseRate(x, reportPurchaseRates),
        "Added Date": formatDate(x.created_at),
      }));
    else
      rows = [
        {
          Sales: s.sales || 0,
          Purchases: s.purchases || 0,
          Returns: s.returns || 0,
          Profit: profit,
        },
      ];
    exportExcel(rows, name.toLowerCase().replaceAll(" ", "-"));
  };
  return (
    <div className="panel-pad">
      <div className="visual-kpis">
        <article>
          <span>Net Sales</span>
          <b>{money((s.sales || 0) - (s.returns || 0))}</b>
          <small>Sales less returns</small>
        </article>
        <article>
          <span>Gross Profit</span>
          <b>{money(profit)}</b>
          <small>Sales minus purchases and returns</small>
        </article>
        <article>
          <span>Account Movement</span>
          <b>{money(s.cash || 0)}</b>
          <small>Money in less money out</small>
        </article>
        <article>
          <span>Payments</span>
          <b>{money(s.payments || 0)}</b>
          <small>Recorded collections</small>
        </article>
      </div>
      <div className="live-financials">
        <article>
          <span>Live Stock Value</span>
          <b>{money(stockValue)}</b>
          <small>Current stock × purchase rate</small>
        </article>
        <article>
          <span>Net Sales</span>
          <b>{money(netSales)}</b>
          <small>Sales less returns</small>
        </article>
        <article>
          <span>Total Purchase Cost</span>
          <b>{money(s.purchases || 0)}</b>
          <small>Imported product value</small>
        </article>
        <article className={profit < 0 ? "loss" : "profit"}>
          <span>Live Profit</span>
          <b>{money(profit)}</b>
          <small>{profitMargin.toFixed(1)}% profit margin</small>
        </article>
      </div>
      <section className="card dealer-report">
        <div className="card-title">
          <h2>Best-Selling Products</h2>
          <small>Ranked by quantity sold from your sales report</small>
        </div>
        <ModuleFilters
          query={productQuery}
          setQuery={setProductQuery}
          placeholder="Search product name or code"
        />
        {visibleBestSellers.length ? (
          <table>
            <thead>
              <tr>
                <th>Rank</th>
                <th>Product Code</th>
                <th>Product Name</th>
                <th>Quantity Sold</th>
                <th>Sales Value</th>
              </tr>
            </thead>
            <tbody>
              {visibleBestSellers.map((x, i) => (
                <tr key={`${x.code}-${x.item}`}>
                  <td>#{i + 1}</td>
                  <td>{String(x.code || "—")}</td>
                  <td>{String(x.item)}</td>
                  <td>{Number(x.quantity).toLocaleString("en-IN")}</td>
                  <td>
                    <b>{money(Number(x.revenue))}</b>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <Empty text="Import sales data to calculate best-selling products." />
        )}
      </section>
      <section className="card dealer-report">
        <div className="card-title">
          <h2>Best Dealers</h2>
          <small>Ranked by purchase value from your sales report</small>
        </div>
        <ModuleFilters
          query={dealerQuery}
          setQuery={setDealerQuery}
          placeholder="Search dealer name or code"
        />
        {visibleDealers.length ? (
          <table>
            <thead>
              <tr>
                <th>Rank</th>
                <th>Dealer Code</th>
                <th>Dealer Name</th>
                <th>Bills</th>
                <th>Purchase Value</th>
              </tr>
            </thead>
            <tbody>
              {visibleDealers.map((x, i) => (
                <tr key={`${x.code}-${x.name}`}>
                  <td>#{i + 1}</td>
                  <td>{String(x.code || "—")}</td>
                  <td>{String(x.name)}</td>
                  <td>{Number(x.bills).toLocaleString("en-IN")}</td>
                  <td>
                    <b>{money(Number(x.purchaseValue))}</b>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <Empty text="Import sales data to calculate the best dealers." />
        )}
      </section>
      <div className="report-charts">
        <article className="card">
          <div className="card-title">
            <h2>Sales vs Purchases</h2>
          </div>
          <div className="bar-chart">
            {[[sales, purchases]].map((x, i) => (
              <div key={i} style={{ width: "35%" }}>
                <span className="sales-bar" style={{ height: `${x[0]}%` }} />
                <span className="purchase-bar" style={{ height: `${x[1]}%` }} />
                <small>All Data</small>
              </div>
            ))}
          </div>
          <div className="legend">
            <span>
              <i className="blue-dot" /> Sales {money(s.sales || 0)}
            </span>
            <span>
              <i className="amber-dot" /> Purchases {money(s.purchases || 0)}
            </span>
          </div>
        </article>
        <article className="card">
          <div className="card-title">
            <h2>Database Summary</h2>
          </div>
          <div className="composition">
            <div className="big-donut">
              <span>
                <b>{money(s.sales || 0)}</b>Total sales
              </span>
            </div>
            <ul>
              <li>
                <i className="blue-dot" />
                Sales <b>{money(s.sales || 0)}</b>
              </li>
              <li>
                <i className="amber-dot" />
                Purchases <b>{money(s.purchases || 0)}</b>
              </li>
              <li>
                <i className="green-dot" />
                Returns <b>{money(s.returns || 0)}</b>
              </li>
            </ul>
          </div>
        </article>
      </div>
      <div className="report-grid">
        {[
          "Sales Register",
          "Purchase Register",
          "Sales Return Register",
          "GST Summary",
          "Cash Book",
          "Bank Book",
          "Outstanding Report",
          "Stock Valuation",
          "Profit Overview",
        ].map((x) => (
          <button key={x} onClick={() => downloadReport(x)}>
            <span>▥</span>
            <b>{x}</b>
            <small>Download Excel from live records</small>
          </button>
        ))}
      </div>
    </div>
  );
}
function SettingsPanel({ notify }: { notify: (s: string) => void }) {
  const data = useResource("settings");
  const [v, setV] = useState({
    business: "",
    gstin: "",
    prefix: "INV-",
    gst: "18",
  });
  useEffect(() => {
    const row = data.rows[0];
    if (row)
      setV({
        business: String(row.business || ""),
        gstin: String(row.gstin || ""),
        prefix: String(row.prefix || "INV-"),
        gst: String(row.gst || "18"),
      });
  }, [data.rows]);
  return (
    <div className="panel-pad">
      <form
        className="settings"
        onSubmit={async (e) => {
          e.preventDefault();
          const row = data.rows[0];
          await request("settings", row ? "PATCH" : "POST", v, row?.id);
          notify("Business settings updated");
          data.reload();
        }}
      >
        <label>
          Business name
          <input
            value={v.business}
            onChange={(e) => setV({ ...v, business: e.target.value })}
            placeholder="Enter business name"
          />
        </label>
        <label>
          GSTIN
          <input
            value={v.gstin}
            onChange={(e) => setV({ ...v, gstin: e.target.value })}
            placeholder="Enter GSTIN"
          />
        </label>
        <label>
          Invoice prefix
          <input
            value={v.prefix}
            onChange={(e) => setV({ ...v, prefix: e.target.value })}
          />
        </label>
        <label>
          Default GST rate
          <select
            value={v.gst}
            onChange={(e) => setV({ ...v, gst: e.target.value })}
          >
            <option value="5">5%</option>
            <option value="12">12%</option>
            <option value="18">18%</option>
            <option value="28">28%</option>
          </select>
        </label>
        <button className="primary">Save Settings</button>
      </form>
    </div>
  );
}
