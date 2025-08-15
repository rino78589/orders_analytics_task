CREATE INDEX IF NOT EXISTS idx_orders_external_id ON orders(external_id);
CREATE INDEX IF NOT EXISTS idx_orders_delivered_at ON orders(delivered_at);
CREATE INDEX IF NOT EXISTS idx_orders_updated_at ON orders(updated_at);
CREATE INDEX IF NOT EXISTS idx_orders_date ON orders(date);
CREATE INDEX IF NOT EXISTS idx_order_items_order_id ON order_items(order_id);
CREATE INDEX IF NOT EXISTS idx_orders_seller_id ON orders(seller_id);

WITH best_orders AS (
    SELECT
        *,
        ROW_NUMBER() OVER (
            PARTITION BY external_id
            ORDER BY
                delivered_at IS NULL,
                delivered_at DESC,
                updated_at DESC
        ) as rn
    FROM orders
    WHERE date >= date('now', '-' || :days || ' days')
)
-- Выбираем только лучшие записи (rn = 1)
, deduplicated_orders AS (
  SELECT
    id,
    external_id,
    date,
    channel,
    seller_id,
    status,
    updated_at,
    delivered_at
  FROM best_orders
  WHERE rn = 1
)
-- Финальная выборка: соединяем дедуплицированные заказы с другими таблицами
SELECT
  o.id as order_id,
  o.date,
  o.channel,
  COALESCE(s.name, 'UNKNOWN_SELLER') as seller,
  o.external_id,
  oi.sku,
  oi.qty,
  oi.revenue,
  oi.cost,
  (oi.revenue - oi.cost) as margin,
  o.status
FROM deduplicated_orders o
JOIN order_items oi ON o.id = oi.order_id
-- LEFT JOIN для обработки отсутствующих продавцов
LEFT JOIN sellers s ON o.seller_id = s.id
ORDER BY o.date DESC, o.external_id, oi.sku;
    