#!/usr/bin/env python3
import argparse
import json
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent
CURRENT = ROOT / "online_product_costs.jsonl"
OUTPUT = ROOT / "product_channel_mappings.jsonl"
PROCESSED = ROOT / "product_channel_processed_basics.json"
PROGRESS = ROOT / "product_channel_fetch_progress.json"
URL = "https://vm-online.benlai.com/erp-product-webapi/ProLhotse/GetListByBasic"
FIELDS = ["ProductSysNo", "ProductID", "ProductBasicSysNo", "WebSiteSysNo", "ChannelSysNo", "SaleChannelSysNo", "Status", "IsDelete"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch-size", type=int, default=500)
    parser.add_argument("--batch-count", type=int, default=1)
    args = parser.parse_args()
    current = {}
    with CURRENT.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                basic = int(item["productBasicSysNo"])
                if basic in current:
                    raise RuntimeError(f"Duplicate pool ProductBasicSysNo: {basic}")
                current[basic] = item["productId"]
    processed = set()
    if PROCESSED.exists():
        processed = set(json.loads(PROCESSED.read_text(encoding="utf-8")))
    existing = set()
    record_count = 0
    if OUTPUT.exists():
        with OUTPUT.open(encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    existing.add(int(json.loads(line)["ProductSysNo"]))
                    record_count += 1
    pending = sorted(set(current) - processed)
    session = requests.Session()
    started = time.time()
    for _ in range(args.batch_count):
        if not pending:
            break
        batch = pending[:args.batch_size]
        response = session.post(URL, data={"productBasicSysNoList": json.dumps(batch)}, timeout=180)
        response.raise_for_status()
        envelope = response.json()
        if envelope.get("Code") != 0:
            raise RuntimeError(envelope.get("Msg"))
        values = envelope.get("Value") or []
        requested = set(batch)
        appended = 0
        with OUTPUT.open("a", encoding="utf-8") as f:
            for value in values:
                basic = int(value["ProductBasicSysNo"])
                product_sysno = int(value["ProductSysNo"])
                if basic not in requested:
                    raise RuntimeError(f"Unexpected ProductBasicSysNo: {basic}")
                if product_sysno in existing:
                    raise RuntimeError(f"Duplicate ProductSysNo: {product_sysno}")
                row = {key: value.get(key) for key in FIELDS}
                row["pool_product_id"] = current[basic]
                f.write(json.dumps(row, ensure_ascii=False) + "\n")
                existing.add(product_sysno)
                appended += 1
        processed.update(requested)
        temp = PROCESSED.with_suffix(".tmp")
        temp.write_text(json.dumps(sorted(processed)), encoding="utf-8")
        temp.replace(PROCESSED)
        record_count += appended
        pending = [x for x in pending if x not in requested]
        stat = {"requested_basics": len(batch), "returned_records": len(values), "new_records": appended, "remaining_basics": len(pending)}
        print(json.dumps(stat, ensure_ascii=False), flush=True)
    active_count = 0
    if OUTPUT.exists():
        with OUTPUT.open(encoding="utf-8") as f:
            for line in f:
                if line.strip() and not json.loads(line)["IsDelete"]:
                    active_count += 1
    result = {
        "pool_products": len(current),
        "processed_basics": len(processed),
        "products_without_mapping": len(processed - set()) if not pending else None,
        "mapping_records": record_count,
        "active_mapping_records": active_count,
        "remaining_basics": len(pending),
        "coverage_complete": not pending and len(processed) == len(current),
        "elapsed_seconds": round(time.time() - started, 2),
        "output": str(OUTPUT),
    }
    PROGRESS.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
