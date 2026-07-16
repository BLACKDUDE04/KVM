type LinkedRow = Record<string, unknown> & { id: number };

type ProductCandidate = {
  code: string;
  name: string;
  purchaseRate: number;
  purchaseDate: string;
  purchaseId: number;
  sellingRate: number;
  salesDate: string;
  salesId: number;
};

const norm = (value: unknown) =>
  String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const text = (value: unknown) => String(value || "").trim();
const number = (value: unknown) => {
  const parsed = Number(value || 0);
  return Number.isFinite(parsed) ? parsed : 0;
};
const missingLabel = (value: unknown, code: unknown, kind: string) => {
  const label = text(value),
    codeText = text(code);
  return (
    !label ||
    /^unknown\b/i.test(label) ||
    (codeText &&
      [norm(codeText), norm(`${kind} ${codeText}`)].includes(norm(label)))
  );
};
const different = (a: unknown, b: unknown) =>
  String(a ?? "") !== String(b ?? "");
const parseItems = (invoice: LinkedRow) => {
  try {
    const parsed = JSON.parse(String(invoice.items_json || "[]"));
    return Array.isArray(parsed) ? (parsed as Record<string, unknown>[]) : [];
  } catch {
    return [];
  }
};
const runStatements = async (
  db: D1Database,
  statements: D1PreparedStatement[],
) => {
  for (let index = 0; index < statements.length; index += 50)
    await db.batch(statements.slice(index, index + 50));
};
const newer = (
  date: string,
  id: number,
  currentDate: string,
  currentId: number,
) => date > currentDate || (date === currentDate && id > currentId);

export async function reconcileLinkedData(db: D1Database) {
  const [
      customerResult,
      supplierResult,
      productResult,
      purchaseResult,
      invoiceResult,
    ] = await db.batch([
      db.prepare("SELECT * FROM customers"),
      db.prepare("SELECT * FROM suppliers"),
      db.prepare("SELECT * FROM products"),
      db.prepare("SELECT * FROM purchases"),
      db.prepare("SELECT * FROM invoices"),
    ]),
    customers = customerResult.results as LinkedRow[],
    suppliers = supplierResult.results as LinkedRow[],
    purchases = purchaseResult.results as LinkedRow[],
    invoices = invoiceResult.results as LinkedRow[];
  let products = productResult.results as LinkedRow[];

  const candidates = new Map<string, ProductCandidate>();
  const candidateFor = (code: unknown, name: unknown) => {
    const sku = text(code),
      key = norm(sku);
    if (!key) return null;
    const current = candidates.get(key) || {
      code: sku,
      name: "",
      purchaseRate: 0,
      purchaseDate: "",
      purchaseId: 0,
      sellingRate: 0,
      salesDate: "",
      salesId: 0,
    };
    if (
      missingLabel(current.name, current.code, "Product") &&
      text(name) &&
      !missingLabel(name, sku, "Product")
    )
      current.name = text(name);
    candidates.set(key, current);
    return current;
  };

  for (const purchase of purchases) {
    const candidate = candidateFor(purchase.item_code, purchase.item);
    if (!candidate) continue;
    const quantity = number(purchase.quantity),
      rate =
        number(purchase.rate) ||
        (quantity ? number(purchase.total) / quantity : 0),
      date = text(purchase.purchase_date);
    if (
      rate > 0 &&
      newer(date, purchase.id, candidate.purchaseDate, candidate.purchaseId)
    ) {
      candidate.purchaseRate = rate;
      candidate.purchaseDate = date;
      candidate.purchaseId = purchase.id;
    }
  }

  for (const invoice of invoices.filter((row) => row.kind !== "quotation")) {
    for (const item of parseItems(invoice)) {
      const candidate = candidateFor(
        item.code || item.item_code,
        item.item || item.name || item.description,
      );
      if (!candidate) continue;
      const quantity = number(item.quantity || item.qty),
        total = number(item.total || item.product_value),
        rate =
          text(item.rate_source) === "purchase"
            ? 0
            : number(item.rate || item.unit_price) ||
              (quantity ? total / quantity : 0),
        date = text(invoice.created_at);
      if (
        rate > 0 &&
        newer(date, invoice.id, candidate.salesDate, candidate.salesId)
      ) {
        candidate.sellingRate = rate;
        candidate.salesDate = date;
        candidate.salesId = invoice.id;
      }
    }
  }

  const productCodes = new Set(products.map((row) => norm(row.sku)));
  const productCreates = [...candidates.entries()]
    .filter(([key]) => !productCodes.has(key))
    .map(([, candidate]) =>
      db
        .prepare(
          "INSERT OR IGNORE INTO products(name,sku,category,stock,reorder_level,purchase_rate,price,gst_rate,created_at) VALUES(?,?,?,0,5,?,?,18,?)",
        )
        .bind(
          candidate.name || `Product ${candidate.code}`,
          candidate.code,
          "Imported",
          candidate.purchaseRate,
          candidate.sellingRate,
          new Date().toISOString(),
        ),
    );
  await runStatements(db, productCreates);
  if (productCreates.length)
    products = (await db.prepare("SELECT * FROM products").all())
      .results as LinkedRow[];

  const customerByCode = new Map(
      customers
        .filter((row) => norm(row.code))
        .map((row) => [norm(row.code), row]),
    ),
    customerByName = new Map(
      customers
        .filter((row) => norm(row.name))
        .map((row) => [norm(row.name), row]),
    ),
    supplierByCode = new Map(
      suppliers
        .filter((row) => norm(row.code))
        .map((row) => [norm(row.code), row]),
    ),
    supplierByName = new Map(
      suppliers
        .filter((row) => norm(row.name))
        .map((row) => [norm(row.name), row]),
    );

  const customerCreates = new Map<string, { name: string; code: string }>();
  for (const invoice of invoices.filter((row) => row.kind !== "quotation")) {
    const code = text(invoice.customer_code),
      name = text(invoice.customer_name),
      key = norm(code) || norm(name);
    if (
      key &&
      !customerByCode.has(norm(code)) &&
      !customerByName.has(norm(name))
    )
      customerCreates.set(key, {
        name: missingLabel(name, code, "Customer")
          ? code || "Unknown Customer"
          : name,
        code,
      });
  }
  await runStatements(
    db,
    [...customerCreates.values()].map((row) =>
      db
        .prepare("INSERT INTO customers(name,code,balance) VALUES(?,?,0)")
        .bind(row.name, row.code),
    ),
  );

  const supplierCreates = new Map<string, { name: string; code: string }>();
  for (const purchase of purchases) {
    const code = text(purchase.supplier_code),
      name = text(purchase.supplier),
      key = norm(code) || norm(name);
    if (
      key &&
      !supplierByCode.has(norm(code)) &&
      !supplierByName.has(norm(name))
    )
      supplierCreates.set(key, {
        name: missingLabel(name, code, "Supplier")
          ? code || "Unknown Supplier"
          : name,
        code,
      });
  }
  await runStatements(
    db,
    [...supplierCreates.values()].map((row) =>
      db
        .prepare("INSERT INTO suppliers(name,code) VALUES(?,?)")
        .bind(row.name, row.code),
    ),
  );

  const refreshedCustomers = customerCreates.size
      ? ((await db.prepare("SELECT * FROM customers").all())
          .results as LinkedRow[])
      : customers,
    refreshedSuppliers = supplierCreates.size
      ? ((await db.prepare("SELECT * FROM suppliers").all())
          .results as LinkedRow[])
      : suppliers,
    refreshedCustomerByCode = new Map(
      refreshedCustomers
        .filter((row) => norm(row.code))
        .map((row) => [norm(row.code), row]),
    ),
    refreshedCustomerByName = new Map(
      refreshedCustomers
        .filter((row) => norm(row.name))
        .map((row) => [norm(row.name), row]),
    ),
    refreshedSupplierByCode = new Map(
      refreshedSuppliers
        .filter((row) => norm(row.code))
        .map((row) => [norm(row.code), row]),
    ),
    refreshedSupplierByName = new Map(
      refreshedSuppliers
        .filter((row) => norm(row.name))
        .map((row) => [norm(row.name), row]),
    ),
    productByCode = new Map(
      products
        .filter((row) => norm(row.sku))
        .map((row) => [norm(row.sku), row]),
    ),
    productByName = new Map(
      products
        .filter((row) => norm(row.name))
        .map((row) => [norm(row.name), row]),
    );

  const purchaseUpdates: D1PreparedStatement[] = [];
  for (const purchase of purchases) {
    const supplier =
        refreshedSupplierByCode.get(norm(purchase.supplier_code)) ||
        refreshedSupplierByName.get(norm(purchase.supplier)),
      product =
        productByCode.get(norm(purchase.item_code)) ||
        productByName.get(norm(purchase.item)),
      quantity = number(purchase.quantity),
      derivedRate = quantity ? number(purchase.total) / quantity : 0,
      rate =
        number(purchase.rate) || derivedRate || number(product?.purchase_rate),
      total = number(purchase.total) || quantity * rate,
      next = {
        supplier:
          supplier &&
          missingLabel(purchase.supplier, purchase.supplier_code, "Supplier")
            ? text(supplier.name)
            : text(purchase.supplier),
        supplier_code: text(purchase.supplier_code) || text(supplier?.code),
        item:
          product && missingLabel(purchase.item, purchase.item_code, "Product")
            ? text(product.name)
            : text(purchase.item),
        item_code: text(purchase.item_code) || text(product?.sku),
        rate,
        total,
      };
    if (
      Object.entries(next).some(([key, value]) =>
        different(purchase[key], value),
      )
    ) {
      purchaseUpdates.push(
        db
          .prepare(
            "UPDATE purchases SET supplier=?,supplier_code=?,item=?,item_code=?,rate=?,total=? WHERE id=?",
          )
          .bind(
            next.supplier,
            next.supplier_code,
            next.item,
            next.item_code,
            next.rate,
            next.total,
            purchase.id,
          ),
      );
      Object.assign(purchase, next);
    }
  }
  await runStatements(db, purchaseUpdates);

  const latestPurchaseRates = new Map<
    string,
    { rate: number; date: string; id: number }
  >();
  for (const purchase of purchases) {
    const quantity = number(purchase.quantity),
      rate =
        number(purchase.rate) ||
        (quantity ? number(purchase.total) / quantity : 0),
      date = text(purchase.purchase_date);
    for (const key of [norm(purchase.item_code), norm(purchase.item)].filter(
      Boolean,
    )) {
      const current = latestPurchaseRates.get(key);
      if (
        rate > 0 &&
        (!current || newer(date, purchase.id, current.date, current.id))
      )
        latestPurchaseRates.set(key, { rate, date, id: purchase.id });
    }
  }

  const invoiceUpdates: D1PreparedStatement[] = [],
    latestSellingRates = new Map<
      string,
      { rate: number; date: string; id: number }
    >();
  for (const invoice of invoices.filter((row) => row.kind !== "quotation")) {
    const customer =
        refreshedCustomerByCode.get(norm(invoice.customer_code)) ||
        refreshedCustomerByName.get(norm(invoice.customer_name)),
      originalItems = parseItems(invoice),
      nextItems = originalItems.map((item) => {
        const product =
            productByCode.get(norm(item.code || item.item_code)) ||
            productByName.get(norm(item.item || item.name || item.description)),
          code = text(item.code || item.item_code) || text(product?.sku),
          name = missingLabel(
            item.item || item.name || item.description,
            code,
            "Product",
          )
            ? text(product?.name) || `Product ${code}`.trim()
            : text(item.item || item.name || item.description),
          quantity = number(item.quantity || item.qty),
          enteredTotal = number(item.total || item.product_value),
          existingRate = number(item.rate || item.unit_price),
          derivedSalesRate =
            !existingRate && quantity && enteredTotal
              ? enteredTotal / quantity
              : 0,
          reportedRate =
            text(item.rate_source) === "purchase"
              ? 0
              : existingRate || derivedSalesRate,
          purchaseRate =
            latestPurchaseRates.get(norm(code))?.rate ||
            latestPurchaseRates.get(norm(name))?.rate ||
            number(product?.purchase_rate),
          rate =
            existingRate ||
            derivedSalesRate ||
            number(product?.price) ||
            purchaseRate,
          total = enteredTotal || quantity * rate,
          rateSource =
            text(item.rate_source) ||
            (existingRate
              ? "sales"
              : derivedSalesRate
                ? "sales-value"
                : number(product?.price)
                  ? "stock"
                  : purchaseRate
                    ? "purchase"
                    : ""),
          key = norm(code) || norm(name),
          current = latestSellingRates.get(key),
          date = text(invoice.created_at);
        if (
          key &&
          reportedRate > 0 &&
          (!current || newer(date, invoice.id, current.date, current.id))
        )
          latestSellingRates.set(key, {
            rate: reportedRate,
            date,
            id: invoice.id,
          });
        return {
          ...item,
          code,
          item: name,
          quantity,
          rate,
          total,
          rate_source: rateSource,
        };
      }),
      itemTotal = nextItems.reduce((sum, item) => sum + number(item.total), 0),
      nextCustomerName =
        customer &&
        missingLabel(invoice.customer_name, invoice.customer_code, "Customer")
          ? text(customer.name)
          : text(invoice.customer_name),
      nextCustomerCode = text(invoice.customer_code) || text(customer?.code),
      nextCustomerId = number(customer?.id) || number(invoice.customer_id),
      nextSubtotal = number(invoice.subtotal) || itemTotal,
      nextAmount = number(invoice.amount) || nextSubtotal,
      nextJson = JSON.stringify(nextItems);
    if (
      number(invoice.customer_id) !== nextCustomerId ||
      different(invoice.customer_name, nextCustomerName) ||
      different(invoice.customer_code, nextCustomerCode) ||
      different(invoice.items_json, nextJson) ||
      different(invoice.subtotal, nextSubtotal) ||
      different(invoice.amount, nextAmount)
    )
      invoiceUpdates.push(
        db
          .prepare(
            "UPDATE invoices SET customer_id=?,customer_name=?,customer_code=?,items_json=?,subtotal=?,amount=? WHERE id=?",
          )
          .bind(
            nextCustomerId || null,
            nextCustomerName || nextCustomerCode || "Unknown Customer",
            nextCustomerCode,
            nextJson,
            nextSubtotal,
            nextAmount,
            invoice.id,
          ),
      );
  }
  await runStatements(db, invoiceUpdates);

  const productUpdates: D1PreparedStatement[] = [];
  for (const product of products) {
    const key = norm(product.sku) || norm(product.name),
      candidate = candidates.get(norm(product.sku)),
      purchaseRate =
        latestPurchaseRates.get(norm(product.sku))?.rate ||
        latestPurchaseRates.get(norm(product.name))?.rate ||
        candidate?.purchaseRate ||
        0,
      sellingRate =
        latestSellingRates.get(norm(product.sku))?.rate ||
        latestSellingRates.get(norm(product.name))?.rate ||
        candidate?.sellingRate ||
        0,
      nextName =
        missingLabel(product.name, product.sku, "Product") && candidate?.name
          ? candidate.name
          : text(product.name),
      nextPurchaseRate = number(product.purchase_rate) || purchaseRate,
      nextPrice = number(product.price) || sellingRate;
    if (
      key &&
      (different(product.name, nextName) ||
        different(product.purchase_rate, nextPurchaseRate) ||
        different(product.price, nextPrice))
    )
      productUpdates.push(
        db
          .prepare(
            "UPDATE products SET name=?,purchase_rate=?,price=? WHERE id=?",
          )
          .bind(nextName, nextPurchaseRate, nextPrice, product.id),
      );
  }
  await runStatements(db, productUpdates);

  return {
    createdProducts: productCreates.length,
    createdCustomers: customerCreates.size,
    createdSuppliers: supplierCreates.size,
    linkedPurchases: purchaseUpdates.length,
    linkedInvoices: invoiceUpdates.length,
    updatedProducts: productUpdates.length,
  };
}
