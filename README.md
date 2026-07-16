# BillFlow Billing Suite

Production-ready billing and business-management web application built with React, Vinext, Cloudflare Workers and D1.

## Features

- GST sales invoices with item-wise quantity, rate and tax calculations
- Editable invoice status and delete controls
- Sales returns and credit notes with Excel import/export
- Customer master with GSTIN, phone and balance
- Product and stock management with reorder alerts
- Excel/CSV import and Excel export for sales invoices, sales returns and purchases
- Payment collection records
- Separate cash-account and bank-account modules with editable ledgers
- Excel/CSV import and Excel export for cash and bank transactions
- Exact-format imports for sales, purchases, customer master, supplier master, stock, cash and bank reports
- Automatic grouping of item-wise sales rows into complete invoices
- Live dashboard and business reports from database records
- Admin-managed users with backend account-disable enforcement
- Responsive desktop and mobile interface
- Automatic empty-database initialization

## Run locally

Requirements: Node.js 22.13 or newer.

```bash
npm install
npm run dev
```

Open `http://localhost:5173` when using a normal Vite environment. The Sites preview workflow uses its own assigned address.

## Build

```bash
npm run build
npm run validate:artifact
```

## Database

The application expects a Cloudflare D1 binding named `DB`. The binding is declared in `.openai/hosting.json`. Database definitions are in `db/schema.ts`, generated migrations are in `drizzle/`, and the API creates missing tables safely on first use.

Generate a new migration after schema changes:

```bash
npm run db:generate
```

## Purchase Excel format

The first worksheet is imported. Supported column names include:

| Required | Optional |
| --- | --- |
| Date | GSTIN |
| Bill No | GST |
| Supplier | Payment Status |
| Item | |
| Qty | |
| Rate | |
| Amount | |

## GitHub upload

Create an empty GitHub repository, then run:

```bash
git remote add github https://github.com/YOUR_USERNAME/billflow-suite.git
git push -u github main
```

Do not commit secrets or production credentials. Sites-managed database resources do not require secrets in this repository.

## Main source files

- `app/page.tsx` — responsive application interface and workflows
- `app/api/data/route.ts` — authenticated CRUD API
- `app/api/summary/route.ts` — live reporting totals
- `db/schema.ts` — D1 relational schema
- `worker/index.ts` — Cloudflare Worker entry
- `app/globals.css` — full responsive styling
