SELECT COUNT(*) AS detail_rows
FROM callback_wmsx_order AS a
JOIN callback_wmsx_order_detail AS b
    ON a.callback_order_code=b.parent_order_code
LEFT JOIN benlai_product.product_basic_info AS pb
    ON b.sku_code COLLATE utf8mb4_0900_ai_ci = pb.ProductID
WHERE a.inventory_time >= '2026-10-01'
  AND a.inventory_time < '2026-10-30'
  AND a.customer_code = 'BLSH'
  AND a.is_deleted = 0
  AND a.order_type IN (
      'CGRK_XZ','CGRK_GYSZS','CGRK_YCD','CGRK',
      'WMSX_ALLOCATION_IN_DIFF',
      'CGTH','CGTH_XZ','CGTH_RMA','CGTH_GYSZS'
  )
