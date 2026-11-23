#!/usr/bin/env python3
import json
import zipfile
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "来福商品池_成本及二次定价_末次采购成本更新_2026-10-30.xlsx"
OUTPUT = ROOT / "来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx"
CURRENT = ROOT / "online_product_costs.jsonl"
MAPPINGS = ROOT / "product_channel_mappings.jsonl"
PRICES = ROOT / "channel_product_price_records.jsonl"
MEDIANS = ROOT / "product_channel_median_costs.jsonl"
LATEST = ROOT / "latest_purchase_costs_2026-10-01_2026-10-30.jsonl"
SPECIAL = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
PURCHASE = "末次采购成本（采购明细）"
OLD_WEIGHTED = "加权成本（costPrice）"
WEIGHTED = "加权成本（分销/C网/非集采中位数）"
OLD_WEIGHTED_HEADER = "加权成本costPrice(元)"
WEIGHTED_HEADER = "加权成本UnitCost中位数(元)"
CALCULATED = {"成本取值口径", "采用成本(元)", "重量分档", "加价率", "二次定价(元)", "定价状态"}


def num(value):
    if value is None or value == "":
        return None
    return float(value)


def round2(value):
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def equal(actual, expected, tolerance=1e-8):
    if expected is None:
        return actual is None
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return actual is not None and isinstance(actual, (int, float)) and abs(float(actual) - float(expected)) <= tolerance
    return actual == expected


def median(values):
    values = sorted(values)
    n = len(values)
    return values[n // 2] if n % 2 else (values[n // 2 - 1] + values[n // 2]) / Decimal(2)


def load_jsonl(path, key):
    result = {}
    with path.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                assert item[key] not in result
                result[item[key]] = item
    return result


def main():
    pool = load_jsonl(CURRENT, "productId")
    median_rows = load_jsonl(MEDIANS, "product_id")
    latest = load_jsonl(LATEST, "product_id")
    prices = {}
    with PRICES.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                prices[int(item["ProductSysNo"])] = Decimal(str(item["UnitCost"]))
    tiers = {pid: {"分销渠道": [], "C网": [], "非集采全渠道": []} for pid in pool}
    with MAPPINGS.open(encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            item = json.loads(line)
            if item["IsDelete"]:
                continue
            pid = item["pool_product_id"]
            row = (int(item["ProductSysNo"]), int(item["WebSiteSysNo"]), prices[int(item["ProductSysNo"])])
            if item["ChannelSysNo"] == 5 and item["SaleChannelSysNo"] == 5:
                tiers[pid]["分销渠道"].append(row)
            if item["SaleChannelSysNo"] == 1:
                tiers[pid]["C网"].append(row)
            if item["SaleChannelSysNo"] != 2:
                tiers[pid]["非集采全渠道"].append(row)
    expected_costs = {}
    source_stats = Counter()
    for pid in pool:
        source = next((name for name in ("分销渠道", "C网", "非集采全渠道") if tiers[pid][name]), None)
        if source is None:
            expected_costs[pid] = None
            continue
        rows = tiers[pid][source]
        costs = [row[2] for row in rows]
        expected_costs[pid] = {
            "selected_source": source,
            "UnitCost": median(costs),
            "selected_record_count": len(rows),
            "selected_website_count": len(set(row[1] for row in rows)),
            "selected_product_sysnos": ",".join(map(str, sorted(row[0] for row in rows))),
            "selected_websites": ",".join(map(str, sorted(set(row[1] for row in rows)))),
            "unit_cost_min": min(costs),
            "unit_cost_max": max(costs),
        }
        source_stats[source] += 1
        source_stats["多记录" if len(rows) > 1 else "单记录"] += 1

    mismatches = []
    for pid, expected in expected_costs.items():
        actual = median_rows[pid]
        if expected is None:
            if actual["UnitCost"] is not None:
                mismatches.append({"product": pid, "issue": "expected no channel cost"})
            continue
        for key, expected_value in expected.items():
            actual_value = actual[key]
            if key in ("UnitCost", "unit_cost_min", "unit_cost_max"):
                ok = Decimal(str(actual_value)) == expected_value
            else:
                ok = actual_value == expected_value
            if not ok:
                mismatches.append({"product": pid, "field": key, "actual": actual_value, "expected": str(expected_value)})

    sf_wb = load_workbook(SOURCE, read_only=True, data_only=False)
    of_wb = load_workbook(OUTPUT, read_only=True, data_only=False)
    ov_wb = load_workbook(OUTPUT, read_only=True, data_only=True)
    assert of_wb.sheetnames == ["汇总", "定价参数", "数据来源", "二次定价", "末次采购明细", "加权成本取价明细"]
    sf = sf_wb["二次定价"]
    of = of_wb["二次定价"]
    ov = ov_wb["二次定价"]
    assert sf.max_row == of.max_row == ov.max_row == 43520
    assert sf.max_column == of.max_column == ov.max_column == 31
    iterators = [sheet.iter_rows(values_only=True) for sheet in (sf, of, ov)]
    sh, oh, ovh = [list(next(it)) for it in iterators]
    assert oh == ovh == [WEIGHTED_HEADER if name == OLD_WEIGHTED_HEADER else name for name in sh]
    si = {name: idx for idx, name in enumerate(sh)}
    oi = {name: idx for idx, name in enumerate(oh)}
    stats = Counter()
    formula_errors = []
    formula_cells = 0
    pool_ids = set()
    for row_no, (src_formula, out_formula, out) in enumerate(zip(*iterators), start=2):
        pid = str(src_formula[si["商品编号"]]).strip()
        pool_ids.add(pid)
        expected = expected_costs[pid]
        weighted = num(expected["UnitCost"]) if expected else None
        purchase = num(src_formula[si["末次采购成本(元)"]])
        for idx, name in enumerate(sh):
            if name == OLD_WEIGHTED_HEADER or name in CALCULATED:
                continue
            if not equal(out_formula[idx], src_formula[idx]):
                mismatches.append({"row": row_no, "product": pid, "field": name, "issue": "unmodified column changed"})
        if not equal(out[oi[WEIGHTED_HEADER]], weighted):
            mismatches.append({"row": row_no, "product": pid, "field": WEIGHTED_HEADER, "actual": out[oi[WEIGHTED_HEADER]], "expected": weighted})
        category = src_formula[si["商品大类"]]
        delivery = src_formula[si["是否产地直发"]]
        weight = num(src_formula[si["净重(g)"]])
        cost_type = PURCHASE if category in SPECIAL else WEIGHTED
        selected = purchase if cost_type == PURCHASE else weighted
        if delivery in ("云仓", "代发"):
            rate, band = 0.02, "不适用"
            stats["云仓／代发商品"] += 1
        elif delivery == "仓发":
            stats["仓发商品"] += 1
            if weight is None:
                rate, band = None, "重量缺失"
            else:
                rate = 0.05 if weight <= 1000 else 0.08
                band = "≤阈值" if weight <= 1000 else ">阈值"
        else:
            rate, band = None, "不适用"
        price = round2(selected * (1 + rate)) if selected is not None and selected > 0 and rate is not None else None
        if cost_type == PURCHASE and purchase is None:
            status = "本期无采购记录"
        elif selected is None or selected <= 0:
            status = "成本缺失"
        elif rate is None:
            status = "发货方式或重量缺失"
        else:
            status = "已定价"
        stats[cost_type] += 1
        stats[status] += 1
        expected_values = {"成本取值口径": cost_type, "采用成本(元)": selected, "重量分档": band, "加价率": rate, "二次定价(元)": price, "定价状态": status}
        for name, expected_value in expected_values.items():
            actual = out[oi[name]]
            if not equal(actual, expected_value):
                mismatches.append({"row": row_no, "product": pid, "field": name, "actual": actual, "expected": expected_value})
            if isinstance(actual, str) and actual.startswith("#"):
                formula_errors.append({"row": row_no, "field": name, "value": actual})
            formula = out_formula[oi[name]]
            if not (isinstance(formula, str) and formula.startswith("=")):
                mismatches.append({"row": row_no, "product": pid, "field": name, "issue": "formula missing"})
            else:
                formula_cells += 1
                expected_formula = src_formula[si[name]].replace(OLD_WEIGHTED, WEIGHTED) if name == "成本取值口径" else src_formula[si[name]]
                if formula != expected_formula:
                    mismatches.append({"row": row_no, "product": pid, "field": name, "issue": "unexpected formula change"})
    assert all(sum(1 for _ in it) == 0 for it in iterators)
    assert pool_ids == set(pool)

    detail = ov_wb["加权成本取价明细"]
    assert detail.max_row == 43520 and detail.max_column == 11
    detail_seen = set()
    for row_no, row in enumerate(detail.iter_rows(min_row=2, values_only=True), start=2):
        pid = row[0]
        detail_seen.add(pid)
        expected = expected_costs[pid]
        if expected is None:
            continue
        checks = [
            (row[1], num(pool[pid]["productBasicSysNo"])), (row[2], expected["selected_source"]),
            (row[3], num(expected["UnitCost"])), (row[4], expected["selected_record_count"]),
            (row[5], expected["selected_website_count"]), (row[6], expected["selected_product_sysnos"]),
            (row[7], expected["selected_websites"]), (row[8], num(expected["unit_cost_min"])), (row[9], num(expected["unit_cost_max"])),
        ]
        for col_idx, (actual, expected_value) in enumerate(checks, start=2):
            if not equal(actual, expected_value):
                mismatches.append({"sheet": "加权成本取价明细", "row": row_no, "product": pid, "column": col_idx, "actual": actual, "expected": expected_value})
    assert detail_seen == set(pool)

    purchase_detail = ov_wb["末次采购明细"]
    assert purchase_detail.max_row == len(latest) + 1 and purchase_detail.max_column == 11
    purchase_seen = set()
    for row in purchase_detail.iter_rows(min_row=2, values_only=True):
        pid = row[6]
        purchase_seen.add(pid)
        expected = latest[pid]
        for idx, key in enumerate(["inventory_time", "callback_sys_no", "callback_order_code", "warehouse_code", "doc_type", "item_code", "product_id", "product_sysno", "qty", "buy_price", "tax_rate"]):
            exp = expected.get(key)
            actual = row[idx]
            ok = equal(actual, num(exp)) if key in ("qty", "buy_price", "tax_rate", "callback_sys_no", "item_code") else actual == exp
            if not ok:
                mismatches.append({"sheet": "末次采购明细", "product": pid, "field": key, "actual": actual, "expected": exp})
    assert purchase_seen == set(latest)

    summary = ov_wb["汇总"]
    summary_values = {summary.cell(r, 1).value: summary.cell(r, 2).value for r in range(7, 19)}
    expected_summary = {
        "商品池商品总数": 43519, "有本期末次采购记录": len(latest), "无本期末次采购记录": 43519 - len(latest),
        "有渠道加权成本": 43519, "无渠道加权成本": 0, "已定价": stats["已定价"],
        "生鲜／水果本期无采购记录": stats["本期无采购记录"], "成本缺失": stats["成本缺失"],
        "云仓／代发商品": stats["云仓／代发商品"], "仓发商品": stats["仓发商品"],
        "末次采购成本口径商品": stats[PURCHASE], "渠道中位数加权成本口径商品": stats[WEIGHTED],
    }
    for key, value in expected_summary.items():
        if not equal(summary_values.get(key), value):
            mismatches.append({"sheet": "汇总", "field": key, "actual": summary_values.get(key), "expected": value})
    for ws in ov_wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for value in row:
                if isinstance(value, str) and value.startswith(("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A")):
                    formula_errors.append({"sheet": ws.title, "value": value})

    with zipfile.ZipFile(SOURCE) as zf:
        source_features = [n for n in zf.namelist() if n.startswith(("xl/charts/", "xl/drawings/", "xl/media/", "xl/pivotTables/"))]
    with zipfile.ZipFile(OUTPUT) as zf:
        output_features = [n for n in zf.namelist() if n.startswith(("xl/charts/", "xl/drawings/", "xl/media/", "xl/pivotTables/"))]
    assert not source_features and not output_features

    result = {
        "rows_checked": len(pool_ids), "main_columns": len(oh), "main_formula_cells_checked": formula_cells,
        "median_products_independently_recalculated": len(expected_costs), "source_selection": dict(source_stats),
        "weighted_detail_rows_checked": len(detail_seen), "purchase_detail_rows_checked": len(purchase_seen),
        "formula_errors": len(formula_errors), "mismatch_count": len(mismatches), "mismatch_samples": mismatches[:20],
        "summary": expected_summary, "file_size_bytes": OUTPUT.stat().st_size,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    sf_wb.close()
    of_wb.close()
    ov_wb.close()
    if mismatches or formula_errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
