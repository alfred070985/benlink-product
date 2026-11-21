#!/usr/bin/env python3
import json
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
BASE = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
FINAL = ROOT / "来福商品池_一件代发收益及入盘建议_2026-10-30.xlsx"
F = 3.0
F5 = 5.0
TARGET = 3.0
MARKUP = 0.02


def round2(v):
    return float(Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def equal(a, e, tolerance=1e-8):
    if e is None:
        return a is None
    if isinstance(e, float):
        return a is not None and abs(float(a) - e) <= tolerance
    return a == e


def decision(current, cost, fulfillment):
    if current is None or cost is None or current <= 0 or cost <= 0:
        return "不建议入盘", None
    profit = current - cost - fulfillment
    if profit >= TARGET:
        return "建议入盘", profit
    if profit > 0:
        return "谨慎入盘", profit
    return "不建议入盘", profit


def main():
    base_wb = load_workbook(BASE, read_only=True, data_only=True)
    final_wb = load_workbook(FINAL, read_only=True, data_only=True)
    base_ws = base_wb["二次定价"]
    final_ws = final_wb["二次定价"]
    base_rows = base_ws.iter_rows(values_only=True)
    final_rows = final_ws.iter_rows(values_only=True)
    base_headers = list(next(base_rows))
    headers = list(next(final_rows))
    bi = {v: i for i, v in enumerate(base_headers)}
    fi = {v: i for i, v in enumerate(headers)}
    counts = Counter()
    mismatches = []
    errors = []
    rows = 0
    reason_counts = Counter()
    for excel_row, (base, out) in enumerate(zip(base_rows, final_rows), start=2):
        rows += 1
        for name in base_headers:
            if not equal(out[fi[name]], base[bi[name]]):
                mismatches.append({"row": excel_row, "field": "原字段:" + name, "actual": out[fi[name]], "expected": base[bi[name]]})
        cost = base[bi["采用成本(元)"]]
        current = base[bi["商品池售价"]]
        cost = float(cost) if cost is not None else None
        current = float(current) if current is not None else None
        second = round2(cost * (1 + MARKUP)) if cost is not None and cost > 0 else None
        profit_second = second - cost - F if second is not None else None
        low_advice, profit_current = decision(current, cost, F)
        high_advice, profit5 = decision(current, cost, F5)
        break_even = cost + F if cost is not None and cost > 0 else None
        target_price = cost + F + TARGET if break_even is not None else None
        target_diff = current - target_price if current is not None and target_price is not None else None
        if cost is None or cost <= 0:
            reason = "采用成本缺失或≤0，无法证明可盈利，不建议入盘。"
        elif current is None or current <= 0:
            reason = "商品池售价缺失或≤0，无法测算收益，不建议入盘。"
        elif profit_current <= 0:
            reason = "现池售价扣除3元履约成本后不盈利，不建议入盘。"
        elif profit_current < TARGET:
            reason = "现池售价可盈利，但低于3元目标单均收益，建议谨慎入盘。"
        elif profit_second is not None and profit_second <= 0:
            reason = "现池售价达到目标收益；但成本加2%会亏损，不能按加2%售价销售。"
        else:
            reason = "现池售价达到3元目标单均收益，且成本加2%仍可盈利。"
        expected = {
            "一件代发履约成本(元/单)": F,
            "一件代发加2%售价(元)": second,
            "加2%单均收益(元)": profit_second,
            "现池售价单均收益(元)": profit_current,
            "保本售价(元)": break_even,
            "目标售价(元)": target_price,
            "现售价与目标价差(元)": target_diff,
            "入盘建议": low_advice,
            "建议原因": reason,
            "5元履约情景收益(元)": profit5,
            "5元情景建议": high_advice,
        }
        for name, value in expected.items():
            actual = out[fi[name]]
            if not equal(actual, value):
                mismatches.append({"row": excel_row, "field": name, "actual": actual, "expected": value})
            if isinstance(actual, str) and actual.startswith("#"):
                errors.append({"row": excel_row, "field": name, "value": actual})
        counts[low_advice] += 1
        counts["5元:" + high_advice] += 1
        if profit_second is not None:
            counts["加2%盈利" if profit_second > 0 else "加2%不盈利"] += 1
        else:
            counts["加2%无法测算"] += 1
        reason_counts[reason] += 1
    extra_base = sum(1 for _ in base_rows)
    extra_final = sum(1 for _ in final_rows)
    summary = final_wb["汇总"]
    summary_actual = {summary.cell(r, 1).value: summary.cell(r, 2).value for r in range(8, 17)}
    summary_expected = {
        "商品池商品总数": rows,
        "建议入盘（3元履约）": counts["建议入盘"],
        "谨慎入盘（3元履约）": counts["谨慎入盘"],
        "不建议入盘（3元履约）": counts["不建议入盘"],
        "成本加2%仍盈利": counts["加2%盈利"],
        "成本加2%不盈利": counts["加2%不盈利"],
        "建议入盘（5元履约）": counts["5元:建议入盘"],
        "谨慎入盘（5元履约）": counts["5元:谨慎入盘"],
        "不建议入盘（5元履约）": counts["5元:不建议入盘"],
    }
    for name, value in summary_expected.items():
        if summary_actual.get(name) != value:
            mismatches.append({"sheet": "汇总", "field": name, "actual": summary_actual.get(name), "expected": value})
    print(json.dumps({
        "rows_checked": rows,
        "columns": len(headers),
        "base_fields_checked_per_row": len(base_headers),
        "new_fields_checked_per_row": 11,
        "extra_base_rows": extra_base,
        "extra_final_rows": extra_final,
        "formula_errors": len(errors),
        "mismatch_count": len(mismatches),
        "mismatch_samples": mismatches[:10],
        "summary": summary_expected,
        "reason_counts": reason_counts,
        "file_size_bytes": FINAL.stat().st_size,
    }, ensure_ascii=False, indent=2))
    base_wb.close()
    final_wb.close()
    if mismatches or errors or extra_base or extra_final:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
