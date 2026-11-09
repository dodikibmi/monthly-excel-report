-- Control totals the Excel report must reproduce exactly (used by test/verify.py).
-- Amounts are rounded per line first, like on an invoice; rounding only the grand total gives a few cents more or less.
SELECT COUNT(DISTINCT o.order_no) AS orders,
       ROUND(SUM(ROUND(CASE o.status WHEN 'RETURNED' THEN -1 ELSE 1 END * l.qty * l.unit_price * (1 - l.discount_pct / 100.0), 2)), 2) AS revenue,
       ROUND(SUM(ROUND(CASE o.status WHEN 'RETURNED' THEN -1 ELSE 1 END * l.qty * (l.unit_price * (1 - l.discount_pct / 100.0) - p.unit_cost), 2)), 2) AS margin
  FROM orders o JOIN order_lines l ON l.order_no = o.order_no JOIN products p ON p.sku = l.sku
 WHERE o.status <> 'CANCELLED' AND o.order_date >= :start AND o.order_date < :end;
