import { getEnv } from "../../../lib/env-context";
import { ensureSchema, getCurrentUser } from "../data/route";

type Row = Record<string, unknown>;
type AuditError = {
  id: string;
  module: string;
  severity: "error" | "warning";
  record: string;
  message: string;
};

const norm = (value: unknown) =>
  String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

export async function GET(request: Request) {
  try {
    await ensureSchema();
    const user = await getCurrentUser(request);
    if (!user || user.role.toLowerCase() === "viewer")
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    const db = getEnv().DB;
    const [
      customers,
      suppliers,
      products,
      invoices,
      purchases,
      accounts,
      transactions,
      expenses,
    ] = await db.batch([
      db.prepare("SELECT * FROM customers"),
      db.prepare("SELECT * FROM suppliers"),
      db.prepare("SELECT * FROM products"),
      db.prepare("SELECT * FROM invoices"),
      db.prepare("SELECT * FROM purchases"),
      db.prepare("SELECT * FROM accounts"),
      db.prepare("SELECT * FROM account_transactions"),
      db.prepare("SELECT * FROM expenses"),
    ]);
    const customerRows = customers.results as Row[],
      supplierRows = suppliers.results as Row[],
      productRows = products.results as Row[],
      invoiceRows = invoices.results as Row[],
      purchaseRows = purchases.results as Row[],
      accountRows = accounts.results as Row[],
      transactionRows = transactions.results as Row[],
      expenseRows = expenses.results as Row[],
      errors: AuditError[] = [];
    const customerIds = new Set(customerRows.map((x) => Number(x.id))),
      supplierCodes = new Set(
        supplierRows.map((x) => norm(x.code)).filter(Boolean),
      ),
      supplierNames = new Set(
        supplierRows.map((x) => norm(x.name)).filter(Boolean),
      ),
      productCodes = new Set(
        productRows.map((x) => norm(x.sku)).filter(Boolean),
      ),
      productNames = new Set(
        productRows.map((x) => norm(x.name)).filter(Boolean),
      ),
      accountIds = new Set(accountRows.map((x) => Number(x.id))),
      bills = [
        ...invoiceRows.map((x) => norm(x.invoice_no)),
        ...purchaseRows.map((x) => norm(x.bill_no)),
      ].filter(Boolean),
      best = new Map<
        string,
        { code: string; item: string; quantity: number; revenue: number }
      >(),
      dealers = new Map<
        string,
        { code: string; name: string; bills: number; purchaseValue: number }
      >();
    for (const invoice of invoiceRows.filter((x) => x.kind !== "quotation")) {
      const record = String(invoice.invoice_no || `Invoice #${invoice.id}`);
      const dealerCode = String(invoice.customer_code || ""),
        dealerName = String(invoice.customer_name || "Unknown Dealer"),
        dealerKey = norm(dealerCode) || norm(dealerName),
        dealer = dealers.get(dealerKey) || {
          code: dealerCode,
          name: dealerName,
          bills: 0,
          purchaseValue: 0,
        };
      dealer.bills += 1;
      dealer.purchaseValue += Number(invoice.amount || 0);
      dealers.set(dealerKey, dealer);
      if (!customerIds.has(Number(invoice.customer_id)))
        errors.push({
          id: `invoice-customer-${invoice.id}`,
          module: "Sales",
          severity: "error",
          record,
          message: "Customer record is missing or not linked.",
        });
      let items: Row[] = [];
      try {
        items = JSON.parse(String(invoice.items_json || "[]"));
      } catch {
        errors.push({
          id: `invoice-json-${invoice.id}`,
          module: "Sales",
          severity: "error",
          record,
          message: "Invoice item data is damaged.",
        });
      }
      const itemTotal = items.reduce(
        (sum, x) =>
          sum +
          Number(
            x.total || Number(x.quantity || x.qty || 0) * Number(x.rate || 0),
          ),
        0,
      );
      if (
        items.length &&
        Math.abs(itemTotal - Number(invoice.subtotal || invoice.amount || 0)) >
          1
      )
        errors.push({
          id: `invoice-total-${invoice.id}`,
          module: "Sales",
          severity: "warning",
          record,
          message: "Saved invoice total does not match its item values.",
        });
      for (const item of items) {
        const code = String(item.code || ""),
          linkedProduct = productRows.find(
            (product) =>
              (code && norm(product.sku) === norm(code)) ||
              (!code &&
                norm(product.name) ===
                  norm(item.item || item.name || item.description)),
          ),
          name = String(
            item.item ||
              item.name ||
              item.description ||
              linkedProduct?.name ||
              `Product ${code}`.trim() ||
              "Unknown item",
          ),
          key = norm(code) || norm(name),
          quantity = Number(item.quantity || item.qty || 0),
          revenue = Number(item.total || quantity * Number(item.rate || 0));
        if (code && !productCodes.has(norm(code)))
          errors.push({
            id: `invoice-product-${invoice.id}-${key}`,
            module: "Sales / Stock",
            severity: "warning",
            record,
            message: `Item code ${code} is not present in Stock.`,
          });
        const current = best.get(key) || {
          code,
          item: name,
          quantity: 0,
          revenue: 0,
        };
        current.quantity += quantity;
        current.revenue += revenue;
        best.set(key, current);
      }
    }
    for (const purchase of purchaseRows) {
      const record = String(purchase.bill_no || `Purchase #${purchase.id}`),
        supplierCode = norm(purchase.supplier_code),
        supplierName = norm(purchase.supplier),
        itemCode = norm(purchase.item_code),
        itemName = norm(purchase.item);
      if (
        (supplierCode && !supplierCodes.has(supplierCode)) ||
        (!supplierCode && supplierName && !supplierNames.has(supplierName))
      )
        errors.push({
          id: `purchase-supplier-${purchase.id}`,
          module: "Purchases / Suppliers",
          severity: "warning",
          record,
          message: "Supplier is not linked to the Supplier master.",
        });
      if (
        (itemCode && !productCodes.has(itemCode)) ||
        (!itemCode && itemName && !productNames.has(itemName))
      )
        errors.push({
          id: `purchase-product-${purchase.id}`,
          module: "Purchases / Stock",
          severity: "warning",
          record,
          message: "Purchased item is not linked to Stock.",
        });
      if (Number(purchase.quantity) < 0 || Number(purchase.total) < 0)
        errors.push({
          id: `purchase-value-${purchase.id}`,
          module: "Purchases",
          severity: "error",
          record,
          message: "Quantity or product value is negative.",
        });
    }
    for (const tx of transactionRows) {
      const record = String(tx.reference || `Transaction #${tx.id}`);
      if (!accountIds.has(Number(tx.account_id)))
        errors.push({
          id: `tx-account-${tx.id}`,
          module: "Cash / Bank",
          severity: "error",
          record,
          message: "Transaction account no longer exists.",
        });
      const reference = norm(tx.reference),
        narration = norm(tx.particulars),
        looksLikeBill = /^(inv|bill|pur|qt)/.test(reference);
      if (
        looksLikeBill &&
        !bills.some(
          (bill) =>
            reference.includes(bill) ||
            bill.includes(reference) ||
            narration.includes(bill),
        )
      )
        errors.push({
          id: `tx-reference-${tx.id}`,
          module: "Cash / Bank",
          severity: "warning",
          record,
          message: "Bill reference does not match any sales or purchase bill.",
        });
      if (Number(tx.amount) <= 0)
        errors.push({
          id: `tx-amount-${tx.id}`,
          module: "Cash / Bank",
          severity: "error",
          record,
          message: "Transaction amount must be greater than zero.",
        });
    }
    for (const expense of expenseRows) {
      const record = String(expense.reference || `Expense #${expense.id}`);
      if (!accountIds.has(Number(expense.account_id)))
        errors.push({
          id: `expense-account-${expense.id}`,
          module: "Expenses / Accounts",
          severity: "error",
          record,
          message: "Expense is not linked to an existing cash or bank account.",
        });
      if (Number(expense.amount) <= 0)
        errors.push({
          id: `expense-amount-${expense.id}`,
          module: "Expenses",
          severity: "error",
          record,
          message: "Expense amount must be greater than zero.",
        });
    }
    const latestPurchaseRates = new Map<
      string,
      { rate: number; date: string; id: number }
    >();
    for (const purchase of purchaseRows) {
      const keys = [norm(purchase.item_code), norm(purchase.item)].filter(
          Boolean,
        ),
        quantity = Number(purchase.quantity || 0),
        rate =
          Number(purchase.rate || 0) ||
          (quantity ? Number(purchase.total || 0) / quantity : 0),
        date = String(purchase.purchase_date || ""),
        id = Number(purchase.id || 0);
      for (const key of keys) {
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
    const stockValue = productRows.reduce((sum, product) => {
      const key = norm(product.sku) || norm(product.name),
        purchaseRate =
          Number(product.purchase_rate || 0) ||
          latestPurchaseRates.get(key)?.rate ||
          0;
      return sum + Number(product.stock || 0) * purchaseRate;
    }, 0);
    const profitItems = new Map<
      string,
      {
        code: string;
        item: string;
        quantity: number;
        revenue: number;
        cost: number;
        profit: number;
        currentStock: number;
        stockValue: number;
      }
    >();
    for (const invoice of invoiceRows.filter((x) => x.kind !== "quotation")) {
      let items: Row[] = [];
      try {
        const parsed = JSON.parse(String(invoice.items_json || "[]"));
        if (Array.isArray(parsed)) items = parsed;
      } catch {}
      for (const line of items) {
        const code = String(line.code || line.item_code || ""),
          enteredName = String(
            line.item || line.name || line.description || "",
          ),
          product = productRows.find(
            (row) =>
              (code && norm(row.sku) === norm(code)) ||
              (!code && norm(row.name) === norm(enteredName)),
          ),
          item = enteredName || String(product?.name || `Product ${code}`),
          key = norm(code) || norm(item),
          quantity = Number(line.quantity || line.qty || 0),
          revenue = Number(
            line.total ||
              line.product_value ||
              quantity * Number(line.rate || line.unit_price || 0),
          ),
          purchaseRate =
            Number(product?.purchase_rate || 0) ||
            latestPurchaseRates.get(norm(code))?.rate ||
            latestPurchaseRates.get(norm(item))?.rate ||
            latestPurchaseRates.get(norm(product?.sku))?.rate ||
            latestPurchaseRates.get(norm(product?.name))?.rate ||
            0,
          cost = quantity * purchaseRate,
          currentStock = Number(product?.stock || 0),
          current = profitItems.get(key) || {
            code: code || String(product?.sku || ""),
            item,
            quantity: 0,
            revenue: 0,
            cost: 0,
            profit: 0,
            currentStock,
            stockValue: currentStock * purchaseRate,
          };
        current.quantity += quantity;
        current.revenue += revenue;
        current.cost += cost;
        current.profit = current.revenue - current.cost;
        current.currentStock = currentStock;
        current.stockValue = currentStock * purchaseRate;
        profitItems.set(key, current);
      }
    }
    return Response.json({
      errors,
      bestSellers: [...best.values()].sort((a, b) => b.quantity - a.quantity),
      bestDealers: [...dealers.values()].sort(
        (a, b) => b.purchaseValue - a.purchaseValue,
      ),
      profitItems: [...profitItems.values()].sort(
        (a, b) => b.profit - a.profit,
      ),
      stockValue,
      checked: {
        sales: invoiceRows.length,
        purchases: purchaseRows.length,
        transactions: transactionRows.length,
        customers: customerRows.length,
        suppliers: supplierRows.length,
        products: productRows.length,
        expenses: expenseRows.length,
      },
      generatedAt: new Date().toISOString(),
    });
  } catch (error) {
    return Response.json(
      {
        error: error instanceof Error ? error.message : "Unable to audit data",
      },
      { status: 500 },
    );
  }
}
