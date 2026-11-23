#!/usr/bin/env python3
"""
95012（禄品供应商模式）新版报价表 独立 QA 校验（2026-11-13 补全版口径）。

校验项：
  1. 列名不含禁用字样（统一 / 上海履约 / 来福商品池）。
  2. 行数 = 4940、列数 = 32。
  3. 不再有「主底稿无此商品」；原 1752 行已补：标准口径行 + 注明未补原因的边界行。
  4. 测算采用成本、二次定价 有效性与「不定价」一致性。
  5. 二次定价口径抽检：二次定价 ≈ 测算采用成本 × (1 + 二次加价率)，容差 0.02。
  6. 来福池成本(元) 有值行数应 = 重合数（3188）。
  7. 补全行的采用成本与生产库源数据逐行一致，且口径与一级类目匹配。
  8. 补全行定价状态符合重合组规则；来福池成本留空。

生成链：generate_95012_pricing.py → complete_95012_basis.py → 本 QA。
"""
import json
import os
import sys
from collections import Counter
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "02_报价表/禄品供应商模式_95012_成本及二次定价_2026-11-12.xlsx"
SRC = REPO / "05_数据与脚本/95012平台核查/核准_1752商品生产库成本_2026-11-13.json"
FRESH = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
CAL_PURCHASE = "采购成本（purchasePrice）"
CAL_WEIGHTED = "加权成本（costPrice）"
EXPECT_ROWS = 4940
EXPECT_OVERLAP = 3188
EXPECT_COLS = 32

wb = openpyxl.load_workbook(OUT, read_only=True, data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
hdr = list(rows[0])
data = rows[1:]
idx = {h: i for i, h in enumerate(hdr)}
print(f"列数: {len(hdr)}  行数: {len(data)}")

ok = True


def fail(msg):
    global ok
    ok = False
    print("  [FAIL]", msg)


def f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# 1) 列名禁用检查
bad = [h for h in hdr if any(k in str(h) for k in ["统一", "上海履约", "来福商品池"])]
if bad:
    fail(f"存在禁用列名: {bad}")
else:
    print("  [OK] 无禁用列名（统一/上海履约/来福商品池）")

# 2) 行数 / 列数
if len(data) != EXPECT_ROWS:
    fail(f"行数 {len(data)} != 期望 {EXPECT_ROWS}")
else:
    print(f"  [OK] 行数 = {EXPECT_ROWS}")
if len(hdr) != EXPECT_COLS:
    fail(f"列数 {len(hdr)} != 期望 {EXPECT_COLS}")
else:
    print(f"  [OK] 列数 = {EXPECT_COLS}")

ci_adopt = idx["测算采用成本(元)"]
ci_2nd = idx["二次定价(元)"]
ci_rate = idx["二次加价率"]
ci_cal = idx["成本取值口径"]
ci_laifu = idx["来福池成本(元)"]
ci_pcode = idx["商品编号"]
ci_cat = idx["一级类目"]
ci_status = idx["定价状态"]

# 3) 补全状态：不再有「主底稿无此商品」，口径枚举与数量正确
payload = json.loads(SRC.read_text(encoding="utf-8"))
prod = {str(x["ProductId"]).strip(): x for x in payload["items"]}
n_miss = sum(1 for r in data if str(r[ci_cal]) == "主底稿无此商品")
if n_miss != 0:
    fail(f"仍有未补全的「主底稿无此商品」行 {n_miss}")
else:
    print("  [OK] 「主底稿无此商品」= 0（1752 行已补全）")

cal_counter = Counter(str(r[ci_cal]) for r in data if str(r[ci_pcode]).strip() in prod)
n_standard = cal_counter[CAL_PURCHASE] + cal_counter[CAL_WEIGHTED]
n_boundary = sum(v for k, v in cal_counter.items() if k.startswith("未补"))
if n_standard + n_boundary != len(prod):
    fail(f"补全行口径数量 {n_standard}+{n_boundary} != 源数据 {len(prod)}")
else:
    print(f"  [OK] 补全行 {len(prod)}：标准口径 {n_standard}（采购 {cal_counter[CAL_PURCHASE]} / 加权 {cal_counter[CAL_WEIGHTED]}）、边界行 {n_boundary}")
if n_boundary != 16:
    fail(f"边界行数 {n_boundary} != 期望 16")
else:
    print(f"  [OK] 边界行 = {n_boundary}（均注明未补原因）")

# 4) 采用成本/二次定价 有效性与「不定价」一致性
ci_cost = idx["成本价"]
ci_mark = idx["被剔除(原)"]
ci_reason = idx["标注原因(原)"]
adopt_vals = [f(r[ci_adopt]) for r in data]
nd_vals = [f(r[ci_2nd]) for r in data]
n_adopt = sum(1 for x in adopt_vals if x is not None and x > 0)
n_nd = sum(1 for x in nd_vals if x is not None and x > 0)
print(f"  测算采用成本 有效: {n_adopt}/{len(data)}  不定价: {len(data)-n_adopt}")
print(f"  二次定价 有效: {n_nd}/{len(data)}  不定价: {len(data)-n_nd}")
bad_inconsist = 0
for r in data:
    a = f(r[ci_adopt]); t = f(r[ci_2nd])
    a_ok = a is not None and a > 0
    t_ok = t is not None and t > 0
    if (not a_ok) and t_ok:
        bad_inconsist += 1
    if not a_ok:
        cost = f(r[ci_cost]); mark = str(r[ci_mark]); reason = r[ci_reason]
        justified = (mark == "是") or (reason not in (None, "")) or (cost is not None and cost <= 0)
        if not justified:
            bad_inconsist += 1
if bad_inconsist > 0:
    fail(f"不定价一致性异常: {bad_inconsist}")
else:
    print("  [OK] 无效行均符合「成本≤0/空→不定价」且有剔除/标注理由")

# 5) 二次定价口径抽检
bad_rate = 0
checked = 0
for r in data:
    c = f(r[ci_adopt]); p = f(r[ci_2nd]); rt = f(r[ci_rate])
    if c and c > 0 and rt is not None and p:
        exp = c * (1 + rt)
        if abs(p - exp) > 0.02:
            bad_rate += 1
        checked += 1
if bad_rate > 0:
    fail(f"二次定价公式不符 {bad_rate}/{checked}")
else:
    print(f"  [OK] 二次定价公式抽检 全部符合 ≈ 成本×(1+加价率) : {checked}/{checked}")

# 6) 来福池成本有值行数 = 重合数
n_lf = sum(1 for r in data if f(r[ci_laifu]) is not None)
if n_lf != EXPECT_OVERLAP:
    fail(f"来福池成本有值行 {n_lf} != 期望重合 {EXPECT_OVERLAP}")
else:
    print(f"  [OK] 来福池成本(元) 有值行 = 重合数 {EXPECT_OVERLAP}")

# 7) 补全行与生产库源数据逐行一致，且口径与类目匹配
bad_src = 0
bad_cat = 0
for r in data:
    p = str(r[ci_pcode]).strip()
    if p not in prod:
        continue
    cal = str(r[ci_cal])
    if cal.startswith("未补"):
        continue
    expected_cal = CAL_PURCHASE if str(r[ci_cat]) in FRESH else CAL_WEIGHTED
    if cal != expected_cal:
        bad_cat += 1
        continue
    field = "PurchasePrice" if cal == CAL_PURCHASE else "CostPrice"
    expected = f(prod[p].get(field))
    actual = f(r[ci_adopt])
    if expected is None or actual is None or abs(actual - expected) > 0.01:
        bad_src += 1
if bad_cat:
    fail(f"补全行口径与一级类目不匹配 {bad_cat}")
elif bad_src:
    fail(f"补全行采用成本与生产库源数据不符 {bad_src}")
else:
    print(f"  [OK] 补全行采用成本与生产库源数据逐行一致（{n_standard} 行），口径与类目匹配")

# 8) 补全行定价状态 / 来福池成本
bad_status = 0
bad_laifu = 0
for r in data:
    p = str(r[ci_pcode]).strip()
    if p not in prod:
        continue
    expected_status = "已定价" if (f(r[ci_2nd]) or 0) > 0 else "成本缺失"
    if str(r[ci_status]) != expected_status:
        bad_status += 1
    if f(r[ci_laifu]) is not None:
        bad_laifu += 1
if bad_status or bad_laifu:
    fail(f"补全行定价状态异常 {bad_status} / 来福池成本未留空 {bad_laifu}")
else:
    print(f"  [OK] 补全行定价状态符合重合组规则，来福池成本留空")

print("\nQA 结论:", "全部通过" if ok else "存在失败项，见上")
sys.exit(0 if ok else 1)
