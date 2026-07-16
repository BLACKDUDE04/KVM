import { getEnv } from "../../../lib/env-context";
import { ensureSchema, getCurrentUser } from "../data/route";

type Row = Record<string, unknown>;

const norm = (value: unknown) =>
  String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
const value = (row: Row | undefined) => Number(row?.value || 0);

export async function GET(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user || user.role.toLowerCase() === "viewer")
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    const db = getEnv().DB;
    const [
      salesResult,
      purchaseTotalResult,
      returnResult,
      paymentResult,
      lowStockResult,
      movementResult,
      productsResult,
      purchasesResult,
      invoicesResult,
      accountsResult,
      transactionsResult,
      expensesResult,
    ] = await db.batch([
      db.prepare(
        "SELECT COALESCE(SUM(amount),0) value FROM invoices WHERE kind='invoice'",
      ),
      db.prepare("SELECT COALESCE(SUM(total),0) value FROM purchases"),
      db.prepare("SELECT COALESCE(SUM(amount),0) value FROM sales_returns"),
      db.prepare("SELECT COALESCE(SUM(amount),0) value FROM payments"),
      db.prepare(
        "SELECT COUNT(*) value FROM products WHERE stock <= reorder_level",
      ),
      db.prepare(
        "SELECT COALESCE(SUM(CASE WHEN direction='in' THEN amount ELSE -amount END),0) value FROM account_transactions",
      ),
      db.prepare("SELECT * FROM products"),
      db.prepare("SELECT * FROM purchases"),
      db.prepare("SELECT * FROM invoices WHERE kind='invoice'"),
      db.prepare("SELECT * FROM accounts"),
      db.prepare("SELECT * FROM account_transactions"),
      db.prepare("SELECT COALESCE(SUM(amount),0) value FROM expenses"),
    ]);

    const products = productsResult.results as Row[],
      purchases = purchasesResult.results as Row[],
      invoices = invoicesResult.results as Row[],
      accounts = accountsResult.results as Row[],
      transactions = transactionsResult.results as Row[],
      expenses = value(expensesResult.results[0] as Row),
      latestPurchaseRates = new Map<
        string,
        { rate: number; date: string; id: number }
      >();
    for (const purchase of purchases) {
      const quantity = Number(purchase.quantity || 0),
        rate =
          Number(purchase.rate || 0) ||
          (quantity ? Number(purchase.total || 0) / quantity : 0),
        date = String(purchase.purchase_date || ""),
        id = Number(purchase.id || 0);
      for (const key of [norm(purchase.item_code), norm(purchase.item)].filter(
        Boolean,
      )) {
        const current = latestPurchaseRates.get(key);
        if (
          rate > 0 &&
          (!current ||
            date > current.date ||
            (date === current.date && id > current.id))
        )
          latestPurchaseRates.set(key, { rate, date, id });
      }
    }
    const productByCode = new Map(
        products.map((product) => [norm(product.sku), product]),
      ),
      productByName = new Map(
        products.map((product) => [norm(product.name), product]),
      ),
      purchaseRateFor = (
        product: Row | undefined,
        code: unknown,
        name: unknown,
      ) =>
        Number(product?.purchase_rate || 0) ||
        latestPurchaseRates.get(norm(code))?.rate ||
        latestPurchaseRates.get(norm(name))?.rate ||
        latestPurchaseRates.get(norm(product?.sku))?.rate ||
        latestPurchaseRates.get(norm(product?.name))?.rate ||
        0;

    let soldRevenue = 0,
      costOfGoodsSold = 0;
    for (const invoice of invoices) {
      let items: Row[] = [];
      try {
        const parsed = JSON.parse(String(invoice.items_json || "[]"));
        if (Array.isArray(parsed)) items = parsed;
      } catch {}
      for (const item of items) {
        const code = item.code || item.item_code,
          name = item.item || item.name || item.description,
          product =
            productByCode.get(norm(code)) || productByName.get(norm(name)),
          quantity = Number(item.quantity || item.qty || 0),
          revenue = Number(
            item.total ||
              item.product_value ||
              quantity * Number(item.rate || item.unit_price || 0),
          );
        soldRevenue += revenue;
        costOfGoodsSold += quantity * purchaseRateFor(product, code, name);
      }
    }

    const closingStockValue = products.reduce(
        (sum, product) =>
          sum +
          Number(product.stock || 0) *
            purchaseRateFor(product, product.sku, product.name),
        0,
      ),
      cashIds = new Set(
        accounts
          .filter((account) => String(account.type).toLowerCase() === "cash")
          .map((account) => Number(account.id)),
      ),
      bankIds = new Set(
        accounts
          .filter((account) => String(account.type).toLowerCase() === "bank")
          .map((account) => Number(account.id)),
      ),
      accountBalance = (ids: Set<number>) =>
        accounts
          .filter((account) => ids.has(Number(account.id)))
          .reduce(
            (sum, account) => sum + Number(account.opening_balance || 0),
            0,
          ) +
        transactions
          .filter((transaction) => ids.has(Number(transaction.account_id)))
          .reduce(
            (sum, transaction) =>
              sum +
              (transaction.direction === "in"
                ? Number(transaction.amount || 0)
                : -Number(transaction.amount || 0)),
            0,
          ),
      returns = value(returnResult.results[0] as Row),
      grossProfit = soldRevenue - returns - costOfGoodsSold,
      netProfit = grossProfit - expenses;

    return Response.json({
      sales: value(salesResult.results[0] as Row),
      purchases: value(purchaseTotalResult.results[0] as Row),
      returns,
      payments: value(paymentResult.results[0] as Row),
      lowStock: value(lowStockResult.results[0] as Row),
      cash: value(movementResult.results[0] as Row),
      cashInHand: accountBalance(cashIds),
      bankBalance: accountBalance(bankIds),
      soldRevenue,
      costOfGoodsSold,
      closingStockValue,
      grossProfit,
      expenses,
      netProfit,
      profit: netProfit,
      inventoryPosition: netProfit + closingStockValue,
    });
  } catch (error) {
    return Response.json(
      {
        error:
          error instanceof Error ? error.message : "Unable to load summary",
      },
      { status: 500 },
    );
  }
}
