#!/usr/bin/env python3
import argparse
import json
import time
from pathlib import Path

from execute_archery_query import execute
from archery_client import ArcheryClient

ROOT = Path(__file__).resolve().parent
MAPPINGS = ROOT / "product_channel_mappings.jsonl"
MAPPING_PROGRESS = ROOT / "product_channel_fetch_progress.json"
OUTPUT = ROOT / "channel_product_price_records.jsonl"
PROCESSED = ROOT / "channel_price_processed_sysnos.json"
PROGRESS = ROOT / "channel_price_fetch_progress.json"
INSTANCE = "[online]P0-sqlserver-master-104-138"
DB_NAME = "benlai"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch-size", type=int, default=5000)
    parser.add_argument("--batch-count", type=int, default=10)
    args = parser.parse_args()
    mapping_progress = json.loads(MAPPING_PROGRESS.read_text(encoding="utf-8"))
    if not mapping_progress.get("coverage_complete"):
        raise RuntimeError("Channel mapping collection is incomplete")
    required = set()
    with MAPPINGS.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                if not item["IsDelete"]:
                    required.add(int(item["ProductSysNo"]))
    existing = {}
    if OUTPUT.exists():
        with OUTPUT.open(encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    item = json.loads(line)
                    existing[int(item["ProductSysNo"])] = item
    processed = set()
    if PROCESSED.exists():
        processed = set(json.loads(PROCESSED.read_text(encoding="utf-8")))
    pending = sorted(required - processed)
    client = ArcheryClient()
    if not client.login():
        raise SystemExit("Archery login failed")
    started = time.time()
    for _ in range(args.batch_count):
        if not pending:
            break
        batch = pending[:args.batch_size]
        sql = f"""SELECT SysNo, ProductSysNo, UnitCost, NoTaxUnitCost, TaxRate, CustomerType, SaleType, IsWholeSale, LastUpdateTime
FROM dbo.Product_Price
WHERE ProductSysNo IN ({','.join(map(str, batch))})"""
        result = execute(client, sql, 0, instance=INSTANCE, db_name=DB_NAME)
        if result.get("status") != 0:
            raise RuntimeError(result.get("msg"))
        data = result.get("data") or {}
        if data.get("error") or data.get("warning"):
            raise RuntimeError(f"error={data.get('error')} warning={data.get('warning')}")
        columns = data.get("column_list") or []
        rows = [dict(zip(columns, values)) for values in data.get("rows") or []]
        requested = set(batch)
        for row in rows:
            sysno = int(row["ProductSysNo"])
            if sysno not in requested:
                raise RuntimeError(f"Unexpected ProductSysNo: {sysno}")
            if sysno in existing:
                raise RuntimeError(f"Duplicate ProductSysNo: {sysno}")
            existing[sysno] = row
        temp = OUTPUT.with_suffix(".tmp")
        with temp.open("w", encoding="utf-8") as f:
            for sysno in sorted(existing):
                f.write(json.dumps(existing[sysno], ensure_ascii=False, default=str) + "\n")
        temp.replace(OUTPUT)
        processed.update(requested)
        processed_temp = PROCESSED.with_suffix(".tmp")
        processed_temp.write_text(json.dumps(sorted(processed)), encoding="utf-8")
        processed_temp.replace(PROCESSED)
        pending = [x for x in pending if x not in requested]
        stat = {"requested": len(batch), "price_records": len(rows), "remaining": len(pending), "query_seconds": data.get("query_time")}
        print(json.dumps(stat, ensure_ascii=False), flush=True)
    result = {
        "active_channel_products": len(required),
        "processed_products": len(processed),
        "price_records": len(existing),
        "products_without_price_record": len(processed - set(existing)),
        "remaining_unprocessed": len(pending),
        "coverage_complete": not pending and processed == required,
        "time_filter": None,
        "elapsed_seconds": round(time.time() - started, 2),
        "output": str(OUTPUT),
    }
    PROGRESS.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
