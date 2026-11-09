-- Fictitious wholesale company "Harbor Supplies". Same shape as a typical ERP sales export.
CREATE TABLE customers (customer_id TEXT PRIMARY KEY, name TEXT NOT NULL, region TEXT NOT NULL, segment TEXT NOT NULL);
CREATE TABLE products  (sku TEXT PRIMARY KEY, name TEXT NOT NULL, category TEXT NOT NULL, unit_cost REAL NOT NULL);
CREATE TABLE orders    (order_no TEXT PRIMARY KEY, order_date TEXT NOT NULL, customer_id TEXT NOT NULL REFERENCES customers,
                        status TEXT NOT NULL CHECK (status IN ('INVOICED','CANCELLED','RETURNED')));
CREATE TABLE order_lines (order_no TEXT NOT NULL REFERENCES orders, line_no INTEGER NOT NULL, sku TEXT NOT NULL REFERENCES products,
                          qty INTEGER NOT NULL, unit_price REAL NOT NULL, discount_pct REAL NOT NULL DEFAULT 0,
                          PRIMARY KEY (order_no, line_no));
