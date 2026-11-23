#!/usr/bin/env python3
import argparse
import json
import time
from pathlib import Path
from execute_archery_query import execute
from archery_client import ArcheryClient

ROOT = Path(__file__).resolve().parent
CURRENT = ROOT / "online_product_costs.jsonl"
OUTPUT = ROOT / "product_price_weighted_costs.jsonl"
PROCESSED = ROOT / "product_price_processed_sysnos.json"
PROGRESS = ROOT / "product_price_fetch_progress.json"
INSTANCE = "[online]P0-sqlserver-master-104-138"
DB_NAME = "benlai"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch-size", type=int, default=5000)
    parser.add_argument("--batch-count", type=int, default=3)
    parser.add_argument("--refresh", action="store_true", help="Discard checkpoints and recollect all current Product_Price records")
    args = parser.parse_args()
    current = {}
    with CURRENT.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                current[int(item["sysNo"])] = item
    existing = {}
    processed = set()
    if not args.refresh:
        if OUTPUT.exists():
            with OUTPUT.open(encoding="utf-8") as f:
                for line in f:
                    if line.strip():
                        item = json.loads(line)
                        existing[int(item["ProductSysNo"])] = item
        if PROCESSED.exists():
            processed = set(json.loads(PROCESSED.read_text(encoding="utf-8")))
    pending = sorted(set(current) - set(existing) - processed)
    client = ArcheryClient()
    if not client.login():
        raise SystemExit("Archery login failed")
    started = time.time()
    stats = []
    for _ in range(args.batch_count):
        if not pending:
            break
        sysnos = pending[:args.batch_size]
        sql = f"""SELECT SysNo, ProductSysNo, UnitCost, NoTaxUnitCost, TaxRate, CustomerType, SaleType, IsWholeSale, LastUpdateTime
FROM dbo.Product_Price
WHERE ProductSysNo >= {min(sysnos)} AND ProductSysNo <= {max(sysnos)}
ORDER BY ProductSysNo"""
        result = execute(client, sql, 0, instance=INSTANCE, db_name=DB_NAME)
        if result.get("status") != 0:
            raise RuntimeError(result.get("msg"))
        data = result.get("data") or {}
        if data.get("error") or data.get("warning"):
            raise RuntimeError(f"error={data.get('error')} warning={data.get('warning')}")
        columns = data.get("column_list") or []
        rows = [dict(zip(columns, values)) for values in data.get("rows") or []]
        returned_ids = [int(row["ProductSysNo"]) for row in rows]
        if len(returned_ids) != len(set(returned_ids)):
            raise RuntimeError("Duplicate ProductSysNo returned")
        requested = set(sysnos)
        matched = requested & set(returned_ids)
        for row in rows:
            sysno = int(row["ProductSysNo"])
            if sysno in requested:
                row["product_id"] = current[sysno]["productId"]
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
        stat = {"requested": len(sysnos), "matched": len(matched), "returned_in_range": len(rows), "query_seconds": data.get("query_time"), "remaining_unprocessed": len(pending)}
        stats.append(stat)
        print(json.dumps(stat, ensure_ascii=False), flush=True)
    no_records = processed - set(existing)
    result = {
        "pool_products": len(current),
        "price_records": len(existing),
        "processed_without_price_record": len(no_records),
        "remaining_unprocessed": len(pending),
        "coverage_complete": not pending and len(existing) + len(no_records) == len(current),
        "elapsed_seconds": round(time.time() - started, 2),
        "output": str(OUTPUT),
    }
    PROGRESS.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
