SELECT PoolType, Status, IsDelete, COUNT(*) AS item_rows,
       COUNT(DISTINCT ProductSysNo) AS product_count,
       COUNT(DISTINCT ProductId) AS product_id_count,
       MAX(LastUpdateTime) AS last_update_time
FROM product_pool_item
GROUP BY PoolType, Status, IsDelete
ORDER BY PoolType, Status, IsDelete;
