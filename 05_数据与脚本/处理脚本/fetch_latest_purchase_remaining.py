#!/usr/bin/env python3
import argparse
import json
import time
from execute_archery_query import execute
from archery_client import ArcheryClient
from fetch_latest_purchase_costs import OUTPUT, PROGRESS, build_sql, load_product_ids

PROCESSED = OUTPUT.with_name("latest_purchase_processed_ids.json")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch-size", type=int, default=5000)
    parser.add_argument("--batch-count", type=int, default=1)
    args = parser.parse_args()
    ids = load_product_ids()
    existing = {}
    if OUTPUT.exists():
        with OUTPUT.open(encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    row = json.loads(line)
                    existing[row["product_id"]] = row
    processed = set()
    if PROCESSED.exists():
        processed = set(json.loads(PROCESSED.read_text(encoding="utf-8")))
    pending = [x for x in ids if x not in existing and x not in processed]
    client = ArcheryClient()
    if not client.login():
        raise SystemExit("Archery login failed")
    started = time.time()
    stats = []
    for batch_no in range(args.batch_count):
        if not pending:
            break
        product_ids = pending[:args.batch_size]
        result = execute(client, build_sql(product_ids), 0)
        if result.get("status") != 0:
            raise RuntimeError(f"Query failed: {result.get('msg')}")
        data = result.get("data") or {}
        if data.get("error") or data.get("warning"):
            raise RuntimeError(f"Query issue: error={data.get('error')} warning={data.get('warning')}")
        columns = data.get("column_list") or []
        rows = [dict(zip(columns, values)) for values in data.get("rows") or []]
        returned_ids = {row["product_id"] for row in rows}
        if len(rows) != len(returned_ids):
            raise RuntimeError("Query returned duplicate products")
        requested_set = set(product_ids)
        if len(rows) == 1000 and len(product_ids) > 1000:
            raise RuntimeError("Result appears truncated at 1000 rows; reduce batch size")
        matched_set = returned_ids & requested_set
        for row in rows:
            row.pop("purchase_rank", None)
            if row["product_id"] in requested_set:
                existing[row["product_id"]] = row
        temp = OUTPUT.with_suffix(".tmp")
        with temp.open("w", encoding="utf-8") as f:
            for product_id in sorted(existing):
                f.write(json.dumps(existing[product_id], ensure_ascii=False, default=str) + "\n")
        temp.replace(OUTPUT)
        processed.update(requested_set)
        processed_temp = PROCESSED.with_suffix(".tmp")
        processed_temp.write_text(json.dumps(sorted(processed), ensure_ascii=False), encoding="utf-8")
        processed_temp.replace(PROCESSED)
        pending = [x for x in pending if x not in requested_set]
        stat = {
            "requested": len(product_ids),
            "matched": len(matched_set),
            "returned_in_range": len(rows),
            "query_seconds": data.get("query_time"),
            "full_sql_suffix": (data.get("full_sql") or "")[-120:],
            "remaining_unprocessed": len(pending),
        }
        stats.append(stat)
        print(json.dumps(stat, ensure_ascii=False), flush=True)
    unmatched_processed = len(processed - set(existing))
    PROGRESS.write_text(json.dumps({
        "total_products": len(ids),
        "collected_products": len(existing),
        "processed_without_purchase_record": unmatched_processed,
        "remaining_unprocessed": len(pending),
        "coverage_complete": not pending and len(existing) + unmatched_processed == len(ids),
        "last_run": stats,
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "collected_products": len(existing),
        "processed_without_purchase_record": unmatched_processed,
        "remaining_unprocessed": len(pending),
        "coverage_complete": not pending and len(existing) + unmatched_processed == len(ids),
        "elapsed_seconds": round(time.time() - started, 2),
        "output": str(OUTPUT),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
