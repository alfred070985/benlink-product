#!/usr/bin/env python3
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "商品池10.29.xlsx"
OUTPUT = ROOT / "online_product_costs.jsonl"
URL = "https://vm-online.benlai.com/business-product-servicehost/ProductApi/GetProductListByIds"
BATCH_SIZE = 500


def read_product_ids():
    wb = load_workbook(INPUT, read_only=True, data_only=True)
    ws = wb["DataMessage"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = headers.index("商品编号")
    ids = []
    seen = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        value = row[idx]
        if value is None:
            continue
        product_id = str(value).strip()
        if product_id and product_id not in seen:
            ids.append(product_id)
            seen.add(product_id)
    wb.close()
    return ids


def fetch_batch(product_ids, attempts=4):
    form = urllib.parse.urlencode({"productIdList": json.dumps(product_ids, ensure_ascii=False)}).encode("utf-8")
    request = urllib.request.Request(
        URL,
        data=form,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    last_error = None
    for attempt in range(attempts):
        try:
            with urllib.request.urlopen(request, timeout=90) as response:
                payload = json.loads(response.read().decode("utf-8"))
            code = payload.get("code", payload.get("Code"))
            if code != 0:
                raise RuntimeError(f"SOA code={code}: {payload.get('msg', payload.get('Msg'))}")
            value = payload.get("value", payload.get("Value"))
            if not isinstance(value, list):
                raise RuntimeError("SOA value is not a list")
            return value
        except Exception as exc:
            last_error = exc
            if attempt + 1 < attempts:
                time.sleep(2 ** attempt)
    raise RuntimeError(f"batch failed after {attempts} attempts: {last_error}")


def main():
    ids = read_product_ids()
    completed = set()
    if OUTPUT.exists():
        with OUTPUT.open("r", encoding="utf-8") as f:
            for line in f:
                try:
                    completed.add(json.loads(line)["productId"])
                except (json.JSONDecodeError, KeyError, TypeError):
                    pass
    remaining = [x for x in ids if x not in completed]
    print(json.dumps({"unique_ids": len(ids), "already_fetched": len(completed), "remaining": len(remaining)}, ensure_ascii=False), flush=True)
    with OUTPUT.open("a", encoding="utf-8") as f:
        for start in range(0, len(remaining), BATCH_SIZE):
            batch = remaining[start:start + BATCH_SIZE]
            items = fetch_batch(batch)
            returned = set()
            for item in items:
                product_id = item.get("productId") or item.get("ProductId")
                if product_id:
                    f.write(json.dumps(item, ensure_ascii=False) + "\n")
                    returned.add(product_id)
            f.flush()
            missing = sorted(set(batch) - returned)
            print(json.dumps({
                "batch_start": start,
                "requested": len(batch),
                "returned": len(items),
                "missing_in_batch": len(missing),
                "missing_sample": missing[:5],
                "total_completed": len(completed) + start + len(batch),
            }, ensure_ascii=False), flush=True)
            time.sleep(0.05)


if __name__ == "__main__":
    main()
