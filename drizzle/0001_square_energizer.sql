CREATE TABLE `account_transactions` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`account_id` integer NOT NULL,
	`direction` text NOT NULL,
	`amount` real NOT NULL,
	`particulars` text NOT NULL,
	`reference` text,
	`transaction_date` text NOT NULL
);
--> statement-breakpoint
CREATE TABLE `accounts` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`name` text NOT NULL,
	`type` text NOT NULL,
	`bank_name` text,
	`account_last4` text,
	`opening_balance` real DEFAULT 0 NOT NULL,
	`active` integer DEFAULT true NOT NULL
);
--> statement-breakpoint
CREATE TABLE `purchases` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`bill_no` text NOT NULL,
	`supplier` text NOT NULL,
	`gstin` text,
	`item` text NOT NULL,
	`quantity` real NOT NULL,
	`rate` real NOT NULL,
	`gst_amount` real DEFAULT 0 NOT NULL,
	`total` real NOT NULL,
	`purchase_date` text NOT NULL,
	`import_batch` text
);
--> statement-breakpoint
CREATE TABLE `sales_returns` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`credit_note_no` text NOT NULL,
	`invoice_id` integer NOT NULL,
	`customer_id` integer NOT NULL,
	`amount` real NOT NULL,
	`reason` text,
	`created_at` text NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `sales_returns_credit_note_no_unique` ON `sales_returns` (`credit_note_no`);