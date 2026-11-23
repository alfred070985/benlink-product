#!/usr/bin/env python3
import json
from collections import Counter
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
FILE = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
FULFILLMENT = 3.0
TARGET_PROFIT = 3.0

wb = load_workbook(FILE, read_only=True, data_only=True)
ws = wb["二次定价"]
rows = ws.iter_rows(values_only=True)
headers = list(next(rows))
ix = {v: i for i, v in enumerate(headers)}
counts = Counter()
profits = []
for row in rows:
    cost = row[ix["采用成本(元)"]]
    current = row[ix["商品池售价"]]
    if cost is None or cost <= 0 or current is None or current <= 0:
        counts["不建议入盘"] += 1
        counts["成本或售价缺失"] += 1
        continue
    current = float(current)
    cost = float(cost)
    second_price = round(cost * 1.02, 2)
    profit_second = second_price - cost - FULFILLMENT
    profit_current = current - cost - FULFILLMENT
    target_price = cost + FULFILLMENT + TARGET_PROFIT
    if profit_current >= TARGET_PROFIT:
        advice = "建议入盘"
    elif profit_current > 0:
        advice = "谨慎入盘"
    else:
        advice = "不建议入盘"
    counts[advice] += 1
    if profit_second > 0:
        counts["成本加2%盈利"] += 1
    else:
        counts["成本加2%不盈利"] += 1
    if current >= target_price:
        counts["现售价达到目标价"] += 1
    profits.append(profit_current)
print(json.dumps({
    "assumptions": {"fulfillment_per_order": FULFILLMENT, "target_profit_per_order": TARGET_PROFIT, "markup": 0.02},
    "counts": counts,
    "priced_products": len(profits),
    "current_profit_min": min(profits),
    "current_profit_max": max(profits),
}, ensure_ascii=False, default=dict, indent=2))
wb.close()
