#!/usr/bin/env python3
"""
95012 报价表：补全 1752 行权威成本口径（2026-11-13，用户确认口径）

两步生成链（顺序不可颠倒）：
  1. python3 generate_95012_pricing.py   # 生成初版，1752 行标注「主底稿无此商品」
  2. python3 complete_95012_basis.py     # 本脚本：只补全这 1752 行

用户确认的三条口径：
  1. 按权威类目逻辑取生产库当前值：生鲜 5 类 → PurchasePrice，其他 → CostPrice；
     改「测算采用成本(元)」，并在原二次定价有效时按原加价率重算「二次定价(元)」。
  2. 三列按重合组规则填：成本取值口径按类目填标准口径名；
     定价状态 = 二次定价>0 ? 已定价 : 成本缺失；来福池成本(元) 留空（这些商品不在来福池）。
  3. 边界行（源表成本≤0，或生产库口径字段≤0）数值保持原状，仅在口径列注明未补原因。

不做的事（保持与既有口径一致）：
  - 不改「成本价」等源表原口径列；
  - 不改毛利额／毛利率／较平台售价差等派生列（重合组本来就有 312 行派生列不自洽）；
  - 不改 32 列列集、不新增任何字段、不动 3188 行重合组。

用法：python3 complete_95012_basis.py [--dry-run]
"""
import argparse
import json
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "02_报价表/禄品供应商模式_95012_成本及二次定价_2026-11-12.xlsx"
SRC = REPO / "05_数据与脚本/95012平台核查/核准_1752商品生产库成本_2026-11-13.json"

FRESH = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
CAL_PURCHASE = "采购成本（purchasePrice）"
CAL_WEIGHTED = "加权成本（costPrice）"
CAL_NOT_FILLED_FIELD = "未补（生产库口径字段≤0，沿用源表原值）"
CAL_NOT_FILLED_ZERO = "未补（源表成本为0，用户确认不恢复定价）"
MISS_LABEL = "主底稿无此商品"


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def price2(cost, rate):
    d = (Decimal(str(cost)) * (Decimal("1") + Decimal(str(rate)))).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    return float(d)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="只统计不写回")
    args = ap.parse_args()

    payload = json.loads(SRC.read_text(encoding="utf-8"))
    prod = {str(x["ProductId"]).strip(): x for x in payload["items"]}
    assert len(prod) == payload["count"] == 1752, "源数据条目数异常"

    wb = openpyxl.load_workbook(OUT, read_only=True, data_only=True)
    sheet_name = wb.active.title
    rows = list(wb.active.iter_rows(values_only=True))
    wb.close()
    hdr = [str(h) for h in rows[0]]
    di = {h: i for i, h in enumerate(hdr)}
    assert len(hdr) == 32, f"列集应为 32 列，实际 {len(hdr)}"
    data = [list(r) for r in rows[1:]]
    assert len(data) == 4940, f"行数应为 4940，实际 {len(data)}"
    original = [list(r) for r in data]

    targets = [i for i, r in enumerate(data) if str(r[di["成本取值口径"]]) == MISS_LABEL]
    if not targets:
        print("已补全（无「主底稿无此商品」行），无需重跑。")
        return
    print(f"目标行数: {len(targets)}（其余 {len(data)-len(targets)} 行保持原样）")

    st = {"filled": 0, "cost_changed": 0, "price_changed": 0, "boundary": 0,
          "boundary_zero": 0, "boundary_field": 0}
    for i in targets:
        r = data[i]
        p = str(r[di["商品编号"]]).strip()
        assert p in prod, f"源数据缺 {p}"
        v = prod[p]
        cat = str(r[di["一级类目"]])
        field = "PurchasePrice" if cat in FRESH else "CostPrice"
        suggested = num(v.get(field))
        old_src = num(r[di["成本价"]])
        old_adopt = num(r[di["测算采用成本(元)"]])
        old_price = num(r[di["二次定价(元)"]])
        rate = num(r[di["二次加价率"]])

        # 边界行：数值一律保持原状，只注明未补原因 + 按重合组规则填定价状态
        if old_src is None or old_src <= 0:
            r[di["成本取值口径"]] = CAL_NOT_FILLED_ZERO
            st["boundary"] += 1
            st["boundary_zero"] += 1
        elif suggested is None or suggested <= 0:
            r[di["成本取值口径"]] = CAL_NOT_FILLED_FIELD
            st["boundary"] += 1
            st["boundary_field"] += 1
        else:
            r[di["成本取值口径"]] = CAL_PURCHASE if cat in FRESH else CAL_WEIGHTED
            if old_adopt is None or abs(old_adopt - suggested) >= 0.01:
                r[di["测算采用成本(元)"]] = suggested
                st["cost_changed"] += 1
            if old_price is None or old_price <= 0 or rate is None:
                raise SystemExit(f"非边界行 {p} 无有效二次定价／加价率，停止（不擅自新增定价）")
            new_price = price2(suggested, rate)
            if abs(new_price - old_price) >= 0.01:
                r[di["二次定价(元)"]] = new_price
                st["price_changed"] += 1
            st["filled"] += 1

        # 三列口径：来福池成本留空；定价状态照重合组规则
        r[di["来福池成本(元)"]] = ""
        r[di["定价状态"]] = "已定价" if (num(r[di["二次定价(元)"]]) or 0) > 0 else "成本缺失"

    # 非目标行必须一字未动
    touched = set(targets)
    for i, (a, b) in enumerate(zip(original, data)):
        if i not in touched and a != b:
            raise SystemExit(f"第 {i+2} 行（非目标行）被意外改动，停止")

    print(f"  按权威口径补全: {st['filled']} 行")
    print(f"  边界行保持原状: {st['boundary']} 行（源表成本为0: {st['boundary_zero']}，生产库口径字段≤0: {st['boundary_field']}）")
    print(f"  改动采用成本: {st['cost_changed']} 行 | 重算二次定价: {st['price_changed']} 行")
    print(f"  口径列: 采购成本 {sum(1 for i in targets if data[i][di['成本取值口径']]==CAL_PURCHASE)} 行 / "
          f"加权成本 {sum(1 for i in targets if data[i][di['成本取值口径']]==CAL_WEIGHTED)} 行")
    print(f"  定价状态: 已定价 {sum(1 for i in targets if data[i][di['定价状态']]=='已定价')} 行 / "
          f"成本缺失 {sum(1 for i in targets if data[i][di['定价状态']]=='成本缺失')} 行")
    print(f"  非目标行未改动: 校验通过 | 列集 {len(hdr)} 列不变")

    if args.dry_run:
        print("  --dry-run：未写回文件")
        return

    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = sheet_name
    ws.append(hdr)
    for r in data:
        ws.append(r)
    wb_out.save(OUT)
    print(f"已写回: {OUT.relative_to(REPO)}")


if __name__ == "__main__":
    main()
