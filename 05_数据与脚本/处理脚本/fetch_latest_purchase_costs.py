#!/usr/bin/env python3
import argparse
import json
import time
from pathlib import Path
from openpyxl import load_workbook
from execute_archery_query import execute
from archery_client import ArcheryClient

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
OUTPUT = ROOT / "latest_purchase_costs_2026-10-01_2026-10-30.jsonl"
PROGRESS = ROOT / "latest_purchase_costs_progress.json"
BATCH_SIZE = 1000

SELECT_COLUMNS = """
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
"""
FROM_WHERE = """
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
"""


def sql_literal(value):
    return "'" + value.replace("\\", "\\\\").replace("'", "''") + "'"


def build_sql(product_ids):
    lower_id = sql_literal(min(product_ids))
    upper_id = sql_literal(max(product_ids))
    # 先在订单明细内取末次记录，再关联商品主数据；结果与原查询口径等价，
    # 但避免对本期全部流水逐行关联商品主数据。
    return f"""SELECT
    inventory_time,
    callback_sys_no,
    callback_order_code,
    warehouse_code,
    doc_type,
    pb.MDMSysNo AS item_code,
    product_id,
    product_sysno,
    qty,
    buy_price,
    CASE WHEN buy_price<>0 AND no_tax_unit_cost<>0
         THEN CAST(buy_price/no_tax_unit_cost-1 AS DECIMAL(4,2))
         ELSE pb.TaxRate END AS tax_rate
FROM (
    SELECT
        a.inventory_time,
        a.callback_sys_no,
        CASE WHEN a.order_type IN ('CGTH','CGTH_XZ','CGTH_RMA','CGTH_GYSZS')
             THEN a.oms_order_code ELSE a.wms_order_code END AS callback_order_code,
        CASE WHEN IFNULL(a.warehouse_code,'')='' AND IFNULL(a.warehouse_id,'')<>''
             THEN a.warehouse_id ELSE a.warehouse_code END AS warehouse_code,
        CASE WHEN a.order_type IN ('CGTH','CGTH_XZ','CGTH_RMA','CGTH_GYSZS')
             THEN '采购退货' ELSE '采购收货' END AS doc_type,
        b.sku_code AS product_id,
        b.channel_goods_code AS product_sysno,
        CASE WHEN a.order_type IN ('CGTH','CGTH_XZ','CGTH_RMA','CGTH_GYSZS')
             THEN -1 ELSE 1 END * real_qty AS qty,
        unit_cost AS buy_price,
        no_tax_unit_cost,
        ROW_NUMBER() OVER (
            PARTITION BY b.sku_code
            ORDER BY a.inventory_time DESC, a.callback_sys_no DESC,
                     a.callback_order_code DESC, b.channel_goods_code DESC,
                     real_qty DESC, unit_cost DESC
        ) AS purchase_rank
    FROM callback_wmsx_order AS a
    JOIN callback_wmsx_order_detail AS b
        ON a.callback_order_code=b.parent_order_code
    WHERE a.inventory_time >= '2026-10-01'
      AND a.inventory_time < '2026-10-30'
      AND a.customer_code = 'BLSH'
      AND a.is_deleted = 0
      AND a.order_type IN (
          'CGRK_XZ','CGRK_GYSZS','CGRK_YCD','CGRK',
          'WMSX_ALLOCATION_IN_DIFF',
          'CGTH','CGTH_XZ','CGTH_RMA','CGTH_GYSZS'
      )
      AND b.sku_code >= {lower_id}
      AND b.sku_code <= {upper_id}
) AS latest_purchase
LEFT JOIN benlai_product.product_basic_info AS pb
    ON product_id COLLATE utf8mb4_0900_ai_ci = pb.ProductID
WHERE purchase_rank = 1
ORDER BY product_id
"""


def load_product_ids():
    wb = load_workbook(INPUT, read_only=True, data_only=True)
    ws = wb["二次定价"]
    it = ws.iter_rows(values_only=True)
    headers = list(next(it))
    idx = headers.index("商品编号")
    ids = sorted({str(row[idx]).strip() for row in it if row[idx] is not None})
    wb.close()
    return ids


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start-batch", type=int, default=0)
    parser.add_argument("--batch-count", type=int, default=10)
    args = parser.parse_args()
    ids = load_product_ids()
    batches = [ids[i:i + BATCH_SIZE] for i in range(0, len(ids), BATCH_SIZE)]
    if args.start_batch < 0 or args.start_batch >= len(batches):
        raise SystemExit("Invalid start batch")
    end = min(args.start_batch + args.batch_count, len(batches))
    client = ArcheryClient()
    if not client.login():
        raise SystemExit("Archery login failed")
    existing = {}
    if OUTPUT.exists():
        with OUTPUT.open(encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    row = json.loads(line)
                    existing[row["product_id"]] = row
    started = time.time()
    stats = []
    for batch_no in range(args.start_batch, end):
        product_ids = batches[batch_no]
        sql = build_sql(product_ids)
        result = execute(client, sql, 1000)
        if result.get("status") != 0:
            raise RuntimeError(f"Batch {batch_no} failed: {result.get('msg')}")
        data = result.get("data") or {}
        if data.get("error") or data.get("warning"):
            raise RuntimeError(f"Batch {batch_no} query issue: error={data.get('error')} warning={data.get('warning')}")
        columns = data.get("column_list") or []
        rows = [dict(zip(columns, values)) for values in data.get("rows") or []]
        returned_ids = {row["product_id"] for row in rows}
        if len(rows) != len(returned_ids):
            raise RuntimeError(f"Batch {batch_no} returned duplicate products")
        if not returned_ids.issubset(set(product_ids)):
            raise RuntimeError(f"Batch {batch_no} returned an unrequested product")
        if len(rows) > len(product_ids):
            raise RuntimeError(f"Batch {batch_no} returned too many rows")
        for row in rows:
            row.pop("purchase_rank", None)
            existing[row["product_id"]] = row
        batch_stat = {
            "batch": batch_no,
            "requested": len(product_ids),
            "matched": len(rows),
            "query_seconds": data.get("query_time"),
        }
        stats.append(batch_stat)
        temp = OUTPUT.with_suffix(".tmp")
        with temp.open("w", encoding="utf-8") as f:
            for product_id in sorted(existing):
                f.write(json.dumps(existing[product_id], ensure_ascii=False, default=str) + "\n")
        temp.replace(OUTPUT)
        PROGRESS.write_text(json.dumps({
            "total_products": len(ids),
            "total_batches": len(batches),
            "last_completed_batch": batch_no,
            "collected_products": len(existing),
            "batches": stats,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(batch_stat, ensure_ascii=False), flush=True)
    print(json.dumps({
        "start_batch": args.start_batch,
        "end_batch_exclusive": end,
        "total_batches": len(batches),
        "collected_products": len(existing),
        "elapsed_seconds": round(time.time() - started, 2),
        "output": str(OUTPUT),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
