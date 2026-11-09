-- One row per invoiced order line for the period, with everything the report needs.
-- :start / :end are the first day of the month and the first day of the next month.
-- Cancelled orders are left out; returned orders count as negative revenue in the month they were returned.
SELECT o.order_no                                   AS "Order",
       o.order_date                                 AS "Date",
       c.name                                       AS "Customer",
       c.region                                     AS "Region",
       c.segment                                    AS "Segment",
       p.category                                   AS "Category",
       p.name                                       AS "Product",
       CASE o.status WHEN 'RETURNED' THEN -l.qty ELSE l.qty END AS "Qty",
       ROUND(CASE o.status WHEN 'RETURNED' THEN -1 ELSE 1 END
             * l.qty * l.unit_price * (1 - l.discount_pct / 100.0), 2) AS "Revenue",
       ROUND(CASE o.status WHEN 'RETURNED' THEN -1 ELSE 1 END
             * l.qty * (l.unit_price * (1 - l.discount_pct / 100.0) - p.unit_cost), 2) AS "Margin"
  FROM orders o
  JOIN order_lines l ON l.order_no = o.order_no
  JOIN customers c   ON c.customer_id = o.customer_id
  JOIN products p    ON p.sku = l.sku
 WHERE o.status <> 'CANCELLED'
   AND o.order_date >= :start AND o.order_date < :end
 ORDER BY o.order_date, o.order_no, l.line_no;
