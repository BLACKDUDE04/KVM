import { getEnv } from "../../../lib/env-context";
import { ensureSchema, getCurrentUser } from "../data/route";

type Row = Record<string, unknown>;
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
    const db = getEnv().DB,
      [
        customersResult,
        invoicesResult,
        returnsResult,
        paymentsResult,
        transactionsResult,
      ] = await db.batch([
        db.prepare("SELECT id, balance FROM customers"),
        db.prepare(
          "SELECT id, invoice_no, customer_id, amount FROM invoices WHERE kind = 'invoice'",
        ),
        db.prepare("SELECT invoice_id, customer_id, amount FROM sales_returns"),
        db.prepare(
          "SELECT invoice_id, customer_id, amount, transaction_id FROM payments",
        ),
        db.prepare(
          "SELECT id, direction, amount, reference, particulars FROM account_transactions WHERE direction = 'in'",
        ),
      ]);
    const customers = customersResult.results as Row[],
      invoices = invoicesResult.results as Row[],
      returns = returnsResult.results as Row[],
      payments = paymentsResult.results as Row[],
      transactions = transactionsResult.results as Row[],
      transactionIds = new Set(transactions.map((row) => Number(row.id))),
      invoiceById = new Map(
        invoices.map((invoice) => [Number(invoice.id), invoice]),
      ),
      linkedPaymentTransactions = new Set<number>(),
      balances = new Map<
        number,
        { opening: number; billed: number; returned: number; paid: number }
      >();
    for (const customer of customers)
      balances.set(Number(customer.id), {
        opening: Number(customer.balance || 0),
        billed: 0,
        returned: 0,
        paid: 0,
      });
    for (const invoice of invoices) {
      const current = balances.get(Number(invoice.customer_id));
      if (current) current.billed += Number(invoice.amount || 0);
    }
    for (const returned of returns) {
      const invoice = invoiceById.get(Number(returned.invoice_id)),
        matchesInvoice =
          invoice &&
          Number(invoice.customer_id) === Number(returned.customer_id),
        current = matchesInvoice
          ? balances.get(Number(returned.customer_id))
          : null;
      if (current) current.returned += Number(returned.amount || 0);
    }
    for (const payment of payments) {
      const transactionId = Number(payment.transaction_id || 0),
        invoice = invoiceById.get(Number(payment.invoice_id)),
        matchesInvoice =
          invoice &&
          Number(invoice.customer_id) === Number(payment.customer_id),
        current = matchesInvoice
          ? balances.get(Number(payment.customer_id))
          : null;
      if (current && transactionId && transactionIds.has(transactionId)) {
        current.paid += Number(payment.amount || 0);
        linkedPaymentTransactions.add(transactionId);
      }
    }
    const invoiceReferences = invoices
      .map((invoice) => ({
        customerId: Number(invoice.customer_id),
        invoiceNo: norm(invoice.invoice_no),
      }))
      .filter((invoice) => invoice.invoiceNo)
      .sort((a, b) => b.invoiceNo.length - a.invoiceNo.length);
    for (const transaction of transactions) {
      if (linkedPaymentTransactions.has(Number(transaction.id))) continue;
      const search = norm(
          `${String(transaction.reference || "")} ${String(transaction.particulars || "")}`,
        ),
        invoice = invoiceReferences.find((item) =>
          search.includes(item.invoiceNo),
        ),
        current = invoice ? balances.get(invoice.customerId) : null;
      if (current) current.paid += Number(transaction.amount || 0);
    }
    return Response.json({
      balances: customers.map((customer) => {
        const detail = balances.get(Number(customer.id))!;
        return {
          customer_id: Number(customer.id),
          ...detail,
          balance:
            detail.opening + detail.billed - detail.returned - detail.paid,
        };
      }),
    });
  } catch (error) {
    return Response.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Unable to calculate customer balances",
      },
      { status: 500 },
    );
  }
}
