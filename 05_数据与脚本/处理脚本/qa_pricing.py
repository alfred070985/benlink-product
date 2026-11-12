#!/usr/bin/env python3
import json
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "商品池10.29.xlsx"
COSTS = ROOT / "online_product_costs.jsonl"
OUTPUT = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
SPECIAL = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
PURCHASE = "采购成本（purchasePrice）"
WEIGHTED = "加权成本（costPrice）"


def num(v):
    if v is None or v == "":
        return None
    return float(v)


def round2(v):
    return float(Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def equal(actual, expected, tolerance=1e-8):
    if expected is None:
        return actual is None
    if isinstance(expected, float):
        return actual is not None and abs(float(actual) - expected) <= tolerance
    return actual == expected


def main():
    costs = {}
    with COSTS.open(encoding="utf-8") as f:
        for line in f:
            item = json.loads(line)
            costs[item["productId"]] = item
    src_wb = load_workbook(INPUT, read_only=True, data_only=True)
    src_ws = src_wb["DataMessage"]
    src_rows = src_ws.iter_rows(values_only=True)
    src_headers = list(next(src_rows))
    si = {v: i for i, v in enumerate(src_headers)}

    out_wb = load_workbook(OUTPUT, read_only=True, data_only=True)
    out_ws = out_wb["二次定价"]
    out_rows = out_ws.iter_rows(values_only=True)
    out_headers = list(next(out_rows))
    oi = {v: i for i, v in enumerate(out_headers)}
    required = src_headers + ["净重(g)", "加权成本costPrice(元)", "采购成本purchasePrice(元)", "成本取值口径", "采用成本(元)", "重量分档", "加价率", "二次定价(元)", "定价状态", "与原池售价差额(元)", "与原池售价差异率"]
    assert all(x in oi for x in required)

    counts = Counter()
    mismatches = []
    formula_errors = []
    rows = 0
    boundary_counts = Counter()
    for excel_row, (src, out) in enumerate(zip(src_rows, out_rows), start=2):
        rows += 1
        pid = str(src[si["商品编号"]]).strip()
        assert out[oi["商品编号"]] == pid
        item = costs[pid]
        category = src[si["商品大类"]]
        delivery = src[si["是否产地直发"]]
        weight = item.get("netWeight")
        weighted = num(item.get("costPrice"))
        purchase = num(item.get("purchasePrice"))
        cost_type = PURCHASE if category in SPECIAL else WEIGHTED
        selected = purchase if cost_type == PURCHASE else weighted
        if delivery in ("云仓", "代发"):
            rate = 0.02
            band = "不适用"
            counts["云仓／代发商品"] += 1
        elif delivery == "仓发":
            counts["仓发商品"] += 1
            if weight is None:
                rate = None
                band = "重量缺失"
            else:
                rate = 0.05 if weight <= 1000 else 0.08
                band = "≤阈值" if weight <= 1000 else ">阈值"
                if weight == 1000:
                    boundary_counts["仓发1000g按5%"] += 1
                elif weight > 1000:
                    boundary_counts["仓发大于1000g按8%"] += 1
                else:
                    boundary_counts["仓发小于1000g按5%"] += 1
        else:
            rate = None
            band = "不适用"
        new_price = round2(selected * (1 + rate)) if selected is not None and selected > 0 and rate is not None else None
        if selected is None or selected <= 0:
            status = "成本缺失"
        elif rate is None:
            status = "发货方式或重量缺失"
        else:
            status = "已定价"
        pool_price = num(src[si["商品池售价"]])
        diff = new_price - pool_price if new_price is not None and pool_price is not None else None
        diff_rate = diff / pool_price if diff is not None and pool_price else None
        expected = {
            "成本取值口径": cost_type,
            "采用成本(元)": selected,
            "重量分档": band,
            "加价率": rate,
            "二次定价(元)": new_price,
            "定价状态": status,
            "与原池售价差额(元)": diff,
            "与原池售价差异率": diff_rate,
        }
        for name, value in expected.items():
            actual = out[oi[name]]
            if not equal(actual, value):
                mismatches.append({"row": excel_row, "product": pid, "field": name, "actual": actual, "expected": value})
        for name in ("二次定价(元)", "与原池售价差额(元)", "与原池售价差异率"):
            value = out[oi[name]]
            if isinstance(value, str) and value.startswith("#"):
                formula_errors.append({"row": excel_row, "field": name, "value": value})
        counts[status] += 1
        counts[cost_type] += 1

    extra_src = sum(1 for _ in src_rows)
    extra_out = sum(1 for _ in out_rows)
    summary = out_wb["汇总"]
    summary_values = {summary.cell(r, 1).value: summary.cell(r, 2).value for r in range(7, 14)}
    expected_summary = {
        "商品池商品总数": rows,
        "已定价": counts["已定价"],
        "成本缺失": counts["成本缺失"],
        "云仓／代发商品": counts["云仓／代发商品"],
        "仓发商品": counts["仓发商品"],
        "采购成本口径商品": counts[PURCHASE],
        "加权成本口径商品": counts[WEIGHTED],
    }
    for key, value in expected_summary.items():
        if summary_values.get(key) != value:
            mismatches.append({"sheet": "汇总", "field": key, "actual": summary_values.get(key), "expected": value})

    print(json.dumps({
        "rows_checked": rows,
        "cost_records": len(costs),
        "columns": len(out_headers),
        "extra_source_rows": extra_src,
        "extra_output_rows": extra_out,
        "formula_cells_expected": rows * 8 + 7,
        "formula_errors": len(formula_errors),
        "mismatch_count": len(mismatches),
        "mismatch_samples": mismatches[:10],
        "summary": expected_summary,
        "warehouse_weight_bands": boundary_counts,
        "file_size_bytes": OUTPUT.stat().st_size,
    }, ensure_ascii=False, indent=2))
    src_wb.close()
    out_wb.close()
    if mismatches or formula_errors or extra_src or extra_out or rows != len(costs):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
