#!/usr/bin/env python3
import json
import zipfile
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
OUTPUT = ROOT / "来福商品池_成本及二次定价_末次采购成本更新_2026-10-30.xlsx"
LATEST = ROOT / "latest_purchase_costs_2026-10-01_2026-10-30.jsonl"
PROCESSED = ROOT / "latest_purchase_processed_ids.json"
PROGRESS = ROOT / "latest_purchase_costs_progress.json"
COUNT_JSON = ROOT / "purchase_callbacks_count.json"
SPECIAL = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
PURCHASE = "末次采购成本（采购明细）"
WEIGHTED = "加权成本（costPrice）"


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


def main():
    latest = {}
    with LATEST.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                assert item["product_id"] not in latest
                latest[item["product_id"]] = item
    processed = set(json.loads(PROCESSED.read_text(encoding="utf-8")))
    no_records = processed - set(latest)
    progress = json.loads(PROGRESS.read_text(encoding="utf-8"))
    raw_detail_rows = json.loads(COUNT_JSON.read_text(encoding="utf-8"))["data"]["rows"][0][0]
    assert progress["coverage_complete"] is True

    sf_wb = load_workbook(SOURCE, read_only=True, data_only=False)
    of_wb = load_workbook(OUTPUT, read_only=True, data_only=False)
    ov_wb = load_workbook(OUTPUT, read_only=True, data_only=True)
    assert of_wb.sheetnames == ["汇总", "定价参数", "数据来源", "二次定价", "末次采购明细"]
    sf = sf_wb["二次定价"]
    of = of_wb["二次定价"]
    ov = ov_wb["二次定价"]
    assert sf.max_row == of.max_row == ov.max_row == 43520
    assert sf.max_column == of.max_column == ov.max_column == 31
    sf_iter = sf.iter_rows(values_only=True)
    of_iter = of.iter_rows(values_only=True)
    ov_iter = ov.iter_rows(values_only=True)
    sh = list(next(sf_iter))
    oh = list(next(of_iter))
    ovh = list(next(ov_iter))
    assert len(sh) == len(oh) == len(ovh) == 31
    assert oh == ovh
    for idx, name in enumerate(sh):
        expected = "末次采购成本(元)" if name == "采购成本purchasePrice(元)" else name
        assert oh[idx] == expected
    si = {name: idx for idx, name in enumerate(sh)}
    oi = {name: idx for idx, name in enumerate(oh)}
    mismatches = []
    formula_errors = []
    stats = Counter()
    pool_ids = set()
    formula_cells = 0

    for row_no, (src, out_formula, out) in enumerate(zip(sf_iter, of_iter, ov_iter), start=2):
        pid = str(src[si["商品编号"]]).strip()
        pool_ids.add(pid)
        item = latest.get(pid)
        purchase = num(item["buy_price"]) if item else None
        for idx, name in enumerate(sh):
            if name == "采购成本purchasePrice(元)":
                continue
            if name in ("成本取值口径", "采用成本(元)", "重量分档", "加价率", "二次定价(元)", "定价状态"):
                continue
            if not equal(out_formula[idx], src[idx]):
                mismatches.append({"row": row_no, "field": name, "issue": "unmodified column changed"})
        if not equal(out[oi["末次采购成本(元)"]], purchase):
            mismatches.append({"row": row_no, "product": pid, "field": "末次采购成本", "actual": out[oi["末次采购成本(元)"]], "expected": purchase})

        category = src[si["商品大类"]]
        delivery = src[si["是否产地直发"]]
        weight = num(src[si["净重(g)"]])
        weighted = num(src[si["加权成本costPrice(元)"]])
        cost_type = PURCHASE if category in SPECIAL else WEIGHTED
        selected = purchase if cost_type == PURCHASE else weighted
        if delivery in ("云仓", "代发"):
            rate = 0.02
            band = "不适用"
            stats["云仓／代发商品"] += 1
        elif delivery == "仓发":
            stats["仓发商品"] += 1
            if weight is None:
                rate = None
                band = "重量缺失"
            else:
                rate = 0.05 if weight <= 1000 else 0.08
                band = "≤阈值" if weight <= 1000 else ">阈值"
        else:
            rate = None
            band = "不适用"
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
        stats["有本期末次采购记录" if item else "无本期末次采购记录"] += 1
        expected = {
            "成本取值口径": cost_type,
            "采用成本(元)": selected,
            "重量分档": band,
            "加价率": rate,
            "二次定价(元)": price,
            "定价状态": status,
        }
        for name, expected_value in expected.items():
            actual = out[oi[name]]
            if not equal(actual, expected_value):
                mismatches.append({"row": row_no, "product": pid, "field": name, "actual": actual, "expected": expected_value})
            if isinstance(actual, str) and actual.startswith("#"):
                formula_errors.append({"row": row_no, "field": name, "value": actual})
        for name in expected:
            value = out_formula[oi[name]]
            if not (isinstance(value, str) and value.startswith("=")):
                mismatches.append({"row": row_no, "field": name, "issue": "formula missing"})
            else:
                formula_cells += 1
        if out_formula[oi["成本取值口径"]] != src[si["成本取值口径"]]:
            mismatches.append({"row": row_no, "field": "成本取值口径", "issue": "lookup formula changed"})
        for name in ("重量分档", "加价率", "二次定价(元)"):
            if out_formula[oi[name]] != src[si[name]]:
                mismatches.append({"row": row_no, "field": name, "issue": "existing formula changed"})
        zref = f"{get_column_letter(oi['成本取值口径'] + 1)}{row_no}"
        xref = f"{get_column_letter(oi['末次采购成本(元)'] + 1)}{row_no}"
        wref = f"{get_column_letter(oi['加权成本costPrice(元)'] + 1)}{row_no}"
        aaref = f"{get_column_letter(oi['采用成本(元)'] + 1)}{row_no}"
        acref = f"{get_column_letter(oi['加价率'] + 1)}{row_no}"
        assert out_formula[oi["采用成本(元)"]] == f'=IF({zref}="{PURCHASE}",{xref},{wref})'
        assert out_formula[oi["定价状态"]] == f'=IF(AND({zref}="{PURCHASE}",{xref}=""),"本期无采购记录",IF(OR({aaref}="",{aaref}<=0),"成本缺失",IF({acref}="","发货方式或重量缺失","已定价")))'

    assert sum(1 for _ in sf_iter) == 0
    assert sum(1 for _ in of_iter) == 0
    assert sum(1 for _ in ov_iter) == 0
    assert len(pool_ids) == 43519
    assert set(latest).issubset(pool_ids)
    assert processed | set(latest) == pool_ids
    assert len(no_records) == 16748

    detail_f = of_wb["末次采购明细"]
    detail_v = ov_wb["末次采购明细"]
    detail_headers = list(next(detail_v.iter_rows(values_only=True)))
    assert detail_f.max_row == detail_v.max_row == len(latest) + 1
    assert detail_f.max_column == detail_v.max_column == 11
    detail_keys = ["inventory_time", "callback_sys_no", "callback_order_code", "warehouse_code", "doc_type", "item_code", "product_id", "product_sysno", "qty", "buy_price", "tax_rate"]
    detail_count = 0
    seen_detail = set()
    for row_no, row in enumerate(detail_v.iter_rows(min_row=2, values_only=True), start=2):
        detail_count += 1
        record = dict(zip(detail_keys, row))
        pid = record["product_id"]
        if pid not in latest or pid in seen_detail:
            mismatches.append({"sheet": "末次采购明细", "row": row_no, "issue": "unknown or duplicate product"})
            continue
        seen_detail.add(pid)
        expected = latest[pid]
        for idx, key in enumerate(detail_keys):
            exp = expected.get(key)
            actual = row[idx]
            if key in ("qty", "buy_price", "tax_rate", "callback_sys_no", "item_code"):
                ok = equal(actual, num(exp))
            else:
                ok = actual == exp
            if not ok:
                mismatches.append({"sheet": "末次采购明细", "row": row_no, "field": key, "actual": actual, "expected": exp})
    assert seen_detail == set(latest)

    summary = ov_wb["汇总"]
    summary_values = {summary.cell(r, 1).value: summary.cell(r, 2).value for r in range(7, 17)}
    expected_summary = {
        "商品池商品总数": 43519,
        "有本期末次采购记录": len(latest),
        "无本期末次采购记录": len(no_records),
        "已定价": stats["已定价"],
        "生鲜／水果本期无采购记录": stats["本期无采购记录"],
        "成本缺失": stats["成本缺失"],
        "云仓／代发商品": stats["云仓／代发商品"],
        "仓发商品": stats["仓发商品"],
        "末次采购成本口径商品": stats[PURCHASE],
        "加权成本口径商品": stats[WEIGHTED],
    }
    for key, value in expected_summary.items():
        if not equal(summary_values.get(key), value):
            mismatches.append({"sheet": "汇总", "field": key, "actual": summary_values.get(key), "expected": value})
    summary_formula_cells = 0
    for row in of_wb["汇总"].iter_rows(values_only=True):
        for value in row:
            if isinstance(value, str) and value.startswith("="):
                summary_formula_cells += 1
    for row in ov_wb["汇总"].iter_rows(values_only=True):
        for value in row:
            if isinstance(value, str) and value.startswith("#"):
                formula_errors.append({"sheet": "汇总", "value": value})

    params = ov_wb["定价参数"]
    param_values = [value for row in params.iter_rows(values_only=True) for value in row if value is not None]
    assert PURCHASE in param_values
    assert "采购成本（purchasePrice）" not in param_values

    with zipfile.ZipFile(SOURCE) as zf:
        source_features = sorted(n for n in zf.namelist() if n.startswith(("xl/charts/", "xl/drawings/", "xl/media/", "xl/pivotTables/")))
    with zipfile.ZipFile(OUTPUT) as zf:
        output_features = sorted(n for n in zf.namelist() if n.startswith(("xl/charts/", "xl/drawings/", "xl/media/", "xl/pivotTables/")))
    assert not source_features and not output_features

    result = {
        "rows_checked": len(pool_ids),
        "main_columns": len(oh),
        "main_formula_cells_checked": formula_cells,
        "summary_formula_cells": summary_formula_cells,
        "raw_purchase_detail_rows": raw_detail_rows,
        "latest_purchase_records_checked": detail_count,
        "products_without_purchase_record_checked": len(no_records),
        "processed_coverage": len(processed | set(latest)),
        "unmodified_columns_preserved": True,
        "detail_fields_checked": detail_count * 11,
        "formula_errors": len(formula_errors),
        "mismatch_count": len(mismatches),
        "mismatch_samples": mismatches[:20],
        "summary": expected_summary,
        "file_size_bytes": OUTPUT.stat().st_size,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    sf_wb.close()
    of_wb.close()
    ov_wb.close()
    if mismatches or formula_errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
