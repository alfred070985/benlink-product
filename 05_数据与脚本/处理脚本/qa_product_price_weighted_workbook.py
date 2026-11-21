#!/usr/bin/env python3
import json
import zipfile
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "来福商品池_成本及二次定价_末次采购成本更新_2026-10-30.xlsx"
OUTPUT = ROOT / "来福商品池_成本及二次定价_Product_Price加权成本_2026-10-30.xlsx"
ERP = ROOT / "product_price_weighted_costs.jsonl"
LATEST = ROOT / "latest_purchase_costs_2026-10-01_2026-10-30.jsonl"
PROCESSED = ROOT / "product_price_processed_sysnos.json"
PROGRESS = ROOT / "product_price_fetch_progress.json"
SPECIAL = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
PURCHASE = "末次采购成本（采购明细）"
OLD_WEIGHTED = "加权成本（costPrice）"
WEIGHTED = "加权成本（Product_Price.UnitCost）"
OLD_WEIGHTED_HEADER = "加权成本costPrice(元)"
WEIGHTED_HEADER = "加权成本UnitCost(元)"
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
    erp = load_jsonl(ERP, "product_id")
    latest = load_jsonl(LATEST, "product_id")
    progress = json.loads(PROGRESS.read_text(encoding="utf-8"))
    assert progress["coverage_complete"] is True
    processed_sysnos = set(json.loads(PROCESSED.read_text(encoding="utf-8")))

    sf_wb = load_workbook(SOURCE, read_only=True, data_only=False)
    sv_wb = load_workbook(SOURCE, read_only=True, data_only=True)
    of_wb = load_workbook(OUTPUT, read_only=True, data_only=False)
    ov_wb = load_workbook(OUTPUT, read_only=True, data_only=True)
    assert of_wb.sheetnames == ["汇总", "定价参数", "数据来源", "二次定价", "末次采购明细"]
    sf = sf_wb["二次定价"]
    sv = sv_wb["二次定价"]
    of = of_wb["二次定价"]
    ov = ov_wb["二次定价"]
    assert sf.max_row == sv.max_row == of.max_row == ov.max_row == 43520
    assert sf.max_column == sv.max_column == of.max_column == ov.max_column == 31
    iterators = [sheet.iter_rows(values_only=True) for sheet in (sf, sv, of, ov)]
    sh, svh, oh, ovh = [list(next(it)) for it in iterators]
    assert sh == svh and oh == ovh
    assert oh == [WEIGHTED_HEADER if name == OLD_WEIGHTED_HEADER else name for name in sh]
    si = {name: idx for idx, name in enumerate(sh)}
    oi = {name: idx for idx, name in enumerate(oh)}
    mismatches = []
    formula_errors = []
    stats = Counter()
    pool_ids = set()
    pool_sysnos = set()
    formula_cells = 0

    for row_no, (src_formula, src, out_formula, out) in enumerate(zip(*iterators), start=2):
        pid = str(src[si["商品编号"]]).strip()
        pool_ids.add(pid)
        pool_sysnos.add(int(src[si["商品系统编号"]]))
        erp_item = erp.get(pid)
        weighted = num(erp_item["UnitCost"]) if erp_item else None
        purchase = num(src[si["末次采购成本(元)"]])
        for idx, name in enumerate(sh):
            if name == OLD_WEIGHTED_HEADER or name in CALCULATED:
                continue
            if not equal(out_formula[idx], src_formula[idx]):
                mismatches.append({"row": row_no, "product": pid, "field": name, "issue": "unmodified column changed"})
        if not equal(out[oi[WEIGHTED_HEADER]], weighted):
            mismatches.append({"row": row_no, "product": pid, "field": WEIGHTED_HEADER, "actual": out[oi[WEIGHTED_HEADER]], "expected": weighted})

        category = src[si["商品大类"]]
        delivery = src[si["是否产地直发"]]
        weight = num(src[si["净重(g)"]])
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
        stats["有Product_Price记录" if erp_item else "无Product_Price记录"] += 1
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
            formula = out_formula[oi[name]]
            if not (isinstance(formula, str) and formula.startswith("=")):
                mismatches.append({"row": row_no, "product": pid, "field": name, "issue": "formula missing"})
            else:
                formula_cells += 1
                expected_formula = src_formula[si[name]]
                if name == "成本取值口径":
                    expected_formula = expected_formula.replace(OLD_WEIGHTED, WEIGHTED)
                if formula != expected_formula:
                    mismatches.append({"row": row_no, "product": pid, "field": name, "issue": "unexpected formula change"})
        if OLD_WEIGHTED in str(out_formula[oi["成本取值口径"]]):
            mismatches.append({"row": row_no, "product": pid, "issue": "old weighted cost type remains"})

    assert all(sum(1 for _ in it) == 0 for it in iterators)
    assert len(pool_ids) == 43519
    assert set(erp).issubset(pool_ids)
    assert set(latest).issubset(pool_ids)
    assert len(erp) == 34671
    assert stats["无Product_Price记录"] == 8848
    assert processed_sysnos == pool_sysnos

    detail = ov_wb["末次采购明细"]
    detail_headers = list(next(detail.iter_rows(values_only=True)))
    detail_keys = ["inventory_time", "callback_sys_no", "callback_order_code", "warehouse_code", "doc_type", "item_code", "product_id", "product_sysno", "qty", "buy_price", "tax_rate"]
    assert detail.max_row == len(latest) + 1 and detail.max_column == 11
    seen_detail = set()
    for row_no, row in enumerate(detail.iter_rows(min_row=2, values_only=True), start=2):
        record = dict(zip(detail_keys, row))
        pid = record["product_id"]
        assert pid in latest and pid not in seen_detail
        seen_detail.add(pid)
        expected = latest[pid]
        for idx, key in enumerate(detail_keys):
            exp = expected.get(key)
            actual = row[idx]
            ok = equal(actual, num(exp)) if key in ("qty", "buy_price", "tax_rate", "callback_sys_no", "item_code") else actual == exp
            if not ok:
                mismatches.append({"sheet": "末次采购明细", "row": row_no, "field": key, "actual": actual, "expected": exp})
    assert seen_detail == set(latest)

    summary = ov_wb["汇总"]
    summary_values = {summary.cell(r, 1).value: summary.cell(r, 2).value for r in range(7, 19)}
    expected_summary = {
        "商品池商品总数": 43519,
        "有本期末次采购记录": len(latest),
        "无本期末次采购记录": 43519 - len(latest),
        "有Product_Price记录": len(erp),
        "无Product_Price记录": stats["无Product_Price记录"],
        "已定价": stats["已定价"],
        "生鲜／水果本期无采购记录": stats["本期无采购记录"],
        "成本缺失": stats["成本缺失"],
        "云仓／代发商品": stats["云仓／代发商品"],
        "仓发商品": stats["仓发商品"],
        "末次采购成本口径商品": stats[PURCHASE],
        "Product_Price加权成本口径商品": stats[WEIGHTED],
    }
    for key, value in expected_summary.items():
        if not equal(summary_values.get(key), value):
            mismatches.append({"sheet": "汇总", "field": key, "actual": summary_values.get(key), "expected": value})
    for row in ov_wb["汇总"].iter_rows(values_only=True):
        for value in row:
            if isinstance(value, str) and value.startswith("#"):
                formula_errors.append({"sheet": "汇总", "value": value})

    params = ov_wb["定价参数"]
    param_values = [value for row in params.iter_rows(values_only=True) for value in row if value is not None]
    assert WEIGHTED in param_values
    assert OLD_WEIGHTED not in param_values
    sources_text = "\n".join(str(value) for row in ov_wb["数据来源"].iter_rows(values_only=True) for value in row if value is not None)
    assert "benlai.dbo.Product_Price" in sources_text and "UnitCost" in sources_text

    with zipfile.ZipFile(SOURCE) as zf:
        source_features = sorted(n for n in zf.namelist() if n.startswith(("xl/charts/", "xl/drawings/", "xl/media/", "xl/pivotTables/")))
    with zipfile.ZipFile(OUTPUT) as zf:
        output_features = sorted(n for n in zf.namelist() if n.startswith(("xl/charts/", "xl/drawings/", "xl/media/", "xl/pivotTables/")))
    assert not source_features and not output_features

    result = {
        "rows_checked": len(pool_ids),
        "main_columns": len(oh),
        "main_formula_cells_checked": formula_cells,
        "product_price_records_checked": len(erp),
        "products_without_product_price_checked": stats["无Product_Price记录"],
        "latest_purchase_records_checked": len(seen_detail),
        "detail_fields_checked": len(seen_detail) * 11,
        "unmodified_columns_preserved": True,
        "formula_errors": len(formula_errors),
        "mismatch_count": len(mismatches),
        "mismatch_samples": mismatches[:20],
        "summary": expected_summary,
        "file_size_bytes": OUTPUT.stat().st_size,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    sf_wb.close()
    sv_wb.close()
    of_wb.close()
    ov_wb.close()
    if mismatches or formula_errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
