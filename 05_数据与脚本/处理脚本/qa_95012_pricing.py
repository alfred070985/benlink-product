#!/usr/bin/env python3
"""
95012（禄品供应商模式）新版报价表 独立 QA 校验。
校验项：
  1. 列名不含禁用字样（统一 / 上海履约 / 来福商品池）。
  2. 行数 = 4940。
  3. 缺失行（主底稿未命中）的新增权威列标注为「主底稿无此商品」。
  4. 测算采用成本、二次定价 非空且 > 0 的比例。
  5. 二次定价口径抽检：二次定价 ≈ 测算采用成本 × (1 + 二次加价率)，容差 0.02。
  6. 来福池成本(元) 有值行数应 = 重合数（3188）。
"""
import openpyxl, sys

OUT = "/Users/duxun/ai-workspace/projects/benlink-product/02_报价表/禄品供应商模式_95012_成本及二次定价_2026-11-12.xlsx"
EXPECT_ROWS = 4940
EXPECT_MISS = 1752
EXPECT_OVERLAP = 3188

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


# 1) 列名禁用检查
bad = [h for h in hdr if any(k in str(h) for k in ["统一", "上海履约", "来福商品池"])]
if bad:
    fail(f"存在禁用列名: {bad}")
else:
    print("  [OK] 无禁用列名（统一/上海履约/来福商品池）")


# 2) 行数
if len(data) != EXPECT_ROWS:
    fail(f"行数 {len(data)} != 期望 {EXPECT_ROWS}")
else:
    print(f"  [OK] 行数 = {EXPECT_ROWS}")


def f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


ci_adopt = idx["测算采用成本(元)"]
ci_2nd = idx["二次定价(元)"]
ci_rate = idx["二次加价率"]
ci_cal = idx["成本取值口径"]
ci_laifu = idx["来福池成本(元)"]


# 3) 缺失标注
n_miss = sum(1 for r in data if str(r[ci_cal]) == "主底稿无此商品")
if n_miss != EXPECT_MISS:
    fail(f"缺失标注行 {n_miss} != 期望 {EXPECT_MISS}")
else:
    print(f"  [OK] 缺失行标注「主底稿无此商品」= {n_miss}")


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
    if (not a_ok) and t_ok:          # 成本无效却定出价 → 矛盾
        bad_inconsist += 1
    if not a_ok:                      # 不定价须有理由
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

print("\nQA 结论:", "全部通过" if ok else "存在失败项，见上")
sys.exit(0 if ok else 1)
