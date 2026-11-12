SELECT
    inventory_time,
    callback_sys_no,
    CASE WHEN a.order_type IN ('CGTH','CGTH_XZ','CGTH_RMA','CGTH_GYSZS')
         THEN a.oms_order_code ELSE a.wms_order_code END AS callback_order_code,
    CASE WHEN IFNULL(a.warehouse_code,'')='' AND IFNULL(a.warehouse_id,'')<>''
         THEN a.warehouse_id ELSE a.warehouse_code END AS warehouse_code,
    CASE WHEN a.order_type IN ('CGTH','CGTH_XZ','CGTH_RMA','CGTH_GYSZS')
         THEN '采购退货' ELSE '采购收货' END AS doc_type,
    pb.MDMSysNo AS item_code,
    b.sku_code AS product_id,
    b.channel_goods_code AS product_sysno,
    CASE WHEN a.order_type IN ('CGTH','CGTH_XZ','CGTH_RMA','CGTH_GYSZS')
         THEN -1 ELSE 1 END * real_qty AS qty,
    unit_cost AS buy_price,
    CASE WHEN unit_cost<>0 AND no_tax_unit_cost<>0
         THEN CAST(unit_cost/no_tax_unit_cost-1 AS DECIMAL(4,2))
         ELSE pb.TaxRate END AS tax_rate
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
