#!/usr/bin/env python3
import json
import statistics
from collections import Counter
from decimal import Decimal
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
CURRENT_API = ROOT / "online_product_costs.jsonl"
ERP = ROOT / "product_price_weighted_costs.jsonl"
PROGRESS = ROOT / "product_price_fetch_progress.json"
BASE = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
DETAIL = ROOT / "product_price_cost_comparison.jsonl"
SUMMARY = ROOT / "product_price_cost_comparison_summary.json"
SPECIAL = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}


def dec(value):
    if value is None or value == "":
        return None
    return Decimal(str(value))


def main():
    progress = json.loads(PROGRESS.read_text(encoding="utf-8"))
    assert progress["coverage_complete"] is True
    current = {}
    with CURRENT_API.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                current[item["productId"]] = item
    erp = {}
    with ERP.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                erp[item["product_id"]] = item
    wb = load_workbook(BASE, read_only=True, data_only=True)
    ws = wb["二次定价"]
    it = ws.iter_rows(values_only=True)
    headers = list(next(it))
    idx = {name: i for i, name in enumerate(headers)}
    stats = Counter()
    weighted_stats = Counter()
    abs_differences = []
    relative_differences = []
    examples = []
    detail_rows = []
    api_excel_mismatches = 0
    for row in it:
        pid = str(row[idx["商品编号"]]).strip()
        category = row[idx["商品大类"]]
        excel_cost = dec(row[idx["加权成本costPrice(元)"]])
        api_item = current[pid]
        api_cost = dec(api_item.get("costPrice"))
        if excel_cost != api_cost:
            api_excel_mismatches += 1
        weighted_scope = category not in SPECIAL
        erp_item = erp.get(pid)
        unit = dec(erp_item["UnitCost"]) if erp_item else None
        no_tax = dec(erp_item["NoTaxUnitCost"]) if erp_item else None
        diff = unit - api_cost if unit is not None and api_cost is not None else None
        relative = diff / api_cost if diff is not None and api_cost != 0 else None
        if erp_item is None:
            status = "无Product_Price记录"
        elif api_cost is None:
            status = "当前costPrice缺失"
        elif api_cost == unit:
            status = "一致"
        else:
            status = "不一致"
        stats[status] += 1
        if weighted_scope:
            weighted_stats[status] += 1
        if erp_item is not None and api_cost is not None:
            if api_cost == no_tax:
                stats["costPrice等于NoTaxUnitCost"] += 1
                if weighted_scope:
                    weighted_stats["costPrice等于NoTaxUnitCost"] += 1
            if api_cost == 0 and unit > 0:
                stats["当前为0_ERP为正"] += 1
                if weighted_scope:
                    weighted_stats["当前为0_ERP为正"] += 1
            elif api_cost > 0 and unit == 0:
                stats["当前为正_ERP为0"] += 1
                if weighted_scope:
                    weighted_stats["当前为正_ERP为0"] += 1
            elif api_cost > 0 and unit > 0 and api_cost != unit:
                stats["双方为正但不一致"] += 1
                if weighted_scope:
                    weighted_stats["双方为正但不一致"] += 1
        if diff is not None:
            abs_diff = abs(diff)
            abs_differences.append(abs_diff)
            if abs_diff <= Decimal("0.000001"):
                stats["差异不超过0.000001"] += 1
            if abs_diff <= Decimal("0.0001"):
                stats["差异不超过0.0001"] += 1
            if abs_diff <= Decimal("0.01"):
                stats["差异不超过0.01"] += 1
            else:
                stats["差异超过0.01"] += 1
            if relative is not None:
                relative_differences.append(abs(relative))
            if status == "不一致":
                examples.append({
                    "product_id": pid,
                    "product_name": api_item.get("productName"),
                    "sysNo": api_item["sysNo"],
                    "category": category,
                    "weighted_scope": weighted_scope,
                    "current_costPrice": str(api_cost),
                    "erp_UnitCost": str(unit),
                    "difference": str(diff),
                    "relative_difference": str(relative) if relative is not None else None,
                    "erp_LastUpdateTime": erp_item.get("LastUpdateTime"),
                })
        detail = {
            "product_id": pid,
            "sysNo": api_item["sysNo"],
            "category": category,
            "weighted_scope": weighted_scope,
            "current_costPrice": str(api_cost) if api_cost is not None else None,
            "erp_UnitCost": str(unit) if unit is not None else None,
            "erp_NoTaxUnitCost": str(no_tax) if no_tax is not None else None,
            "difference_erp_minus_current": str(diff) if diff is not None else None,
            "status": status,
            "erp_LastUpdateTime": erp_item.get("LastUpdateTime") if erp_item else None,
        }
        detail_rows.append(detail)
    wb.close()
    assert len(detail_rows) == len(current) == 43519
    with DETAIL.open("w", encoding="utf-8") as f:
        for item in detail_rows:
            f.write(json.dumps(item, ensure_ascii=False) + "\n")
    examples.sort(key=lambda x: abs(Decimal(x["difference"])), reverse=True)
    matched = stats["一致"] + stats["不一致"]
    result = {
        "pool_products": len(detail_rows),
        "weighted_scope_products": sum(weighted_stats[s] for s in ("一致", "不一致", "无Product_Price记录", "当前costPrice缺失")),
        "product_price_records": len(erp),
        "products_without_product_price": stats["无Product_Price记录"],
        "current_api_vs_excel_mismatches": api_excel_mismatches,
        "all_products": dict(stats),
        "weighted_scope": dict(weighted_stats),
        "match_rate_among_product_price_records": float(Decimal(stats["一致"]) / Decimal(matched)) if matched else None,
        "match_rate_all_pool_products": float(Decimal(stats["一致"]) / Decimal(len(detail_rows))),
        "absolute_difference_median": str(statistics.median(abs_differences)) if abs_differences else None,
        "absolute_difference_max": str(max(abs_differences)) if abs_differences else None,
        "relative_difference_median": str(statistics.median(relative_differences)) if relative_differences else None,
        "top_difference_examples": examples[:20],
        "detail_file": str(DETAIL),
    }
    SUMMARY.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
