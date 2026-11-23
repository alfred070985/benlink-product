#!/usr/bin/env python3
import json
from pathlib import Path
root = Path(__file__).resolve().parent
latest = set()
with (root / "latest_purchase_costs_2026-10-01_2026-10-30.jsonl").open(encoding="utf-8") as f:
    for line in f:
        if line.strip():
            latest.add(json.loads(line)["product_id"])
processed = set(json.loads((root / "latest_purchase_processed_ids.json").read_text(encoding="utf-8")))
intersection = processed & latest
print(json.dumps({"latest": len(latest), "processed": len(processed), "union": len(processed | latest), "intersection": len(intersection), "samples": sorted(intersection)[:30]}, ensure_ascii=False, indent=2))
