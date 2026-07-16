ALTER TABLE `invoices` ADD COLUMN `customer_name` text NOT NULL DEFAULT '';--> statement-breakpoint
ALTER TABLE `invoices` ADD COLUMN `items_json` text NOT NULL DEFAULT '[]';--> statement-breakpoint
ALTER TABLE `invoices` ADD COLUMN `subtotal` real NOT NULL DEFAULT 0;--> statement-breakpoint
ALTER TABLE `invoices` ADD COLUMN `tax` real NOT NULL DEFAULT 0;--> statement-breakpoint
ALTER TABLE `products` ADD COLUMN `category` text NOT NULL DEFAULT 'General';
