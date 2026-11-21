#!/usr/bin/env python3
import json
from collections import Counter
from decimal import Decimal
from pathlib import Path

root = Path(__file__).resolve().parent
current = {}
with (root / "online_product_costs.jsonl").open(encoding="utf-8") as f:
    for line in f:
        if line.strip():
            item = json.loads(line)
            current[item["productId"]] = item
stats = Counter()
with (root / "product_price_cost_comparison.jsonl").open(encoding="utf-8") as f:
    for line in f:
        if not line.strip():
            continue
        item = json.loads(line)
        api = current[item["product_id"]]
        delivery = "仓发" if api.get("deliveryMethod") == 1 else "代发／云仓" if api.get("deliveryMethod") == 2 else "未知"
        scope = "加权成本口径" if item["weighted_scope"] else "末次采购成本口径"
        status = item["status"]
        stats[(delivery, "商品数")] += 1
        stats[(delivery, status)] += 1
        stats[(delivery, scope, status)] += 1
        cost = api.get("costPrice")
        purchase = api.get("purchasePrice")
        if cost is not None and purchase is not None and Decimal(str(cost)) == Decimal(str(purchase)):
            stats[(delivery, "costPrice等于purchasePrice")] += 1
        if status in ("一致", "不一致"):
            stats[(delivery, "有ERP成本记录")] += 1
            if item["weighted_scope"]:
                stats[(delivery, "加权口径有ERP成本记录")] += 1

result = {}
for delivery in ("仓发", "代发／云仓", "未知"):
    total = stats[(delivery, "商品数")]
    if not total:
        continue
    matched_records = stats[(delivery, "有ERP成本记录")]
    result[delivery] = {
        "商品数": total,
        "一致": stats[(delivery, "一致")],
        "不一致": stats[(delivery, "不一致")],
        "无Product_Price记录": stats[(delivery, "无Product_Price记录")],
        "当前costPrice缺失": stats[(delivery, "当前costPrice缺失")],
        "costPrice等于purchasePrice": stats[(delivery, "costPrice等于purchasePrice")],
        "有ERP记录的一致率": stats[(delivery, "一致")] / matched_records if matched_records else None,
        "加权口径_一致": stats[(delivery, "加权成本口径", "一致")],
        "加权口径_不一致": stats[(delivery, "加权成本口径", "不一致")],
        "加权口径_无Product_Price记录": stats[(delivery, "加权成本口径", "无Product_Price记录")],
    }
print(json.dumps(result, ensure_ascii=False, indent=2))
