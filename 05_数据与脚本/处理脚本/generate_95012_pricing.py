#!/usr/bin/env python3
"""
生成 95012（禄品供应商模式）新版报价表 —— 合并补全方案

数据源与口径：
  - 骨架：桌面禄品筛选表（4940 行）「原」口径列，剔除「统一采购成本」实验列。
  - 重合（主底稿命中）的 3188 个：采用成本 / 二次定价 / 二次加价率 用主底稿权威值覆盖；
    并补入主底稿独有且体现权威口径的列：成本取值口径、来福池成本(元)、定价状态。
  - 缺失（主底稿未命中）的 1752 个：沿用桌面「原」口径，新增权威列留空（标注"主底稿无此商品"）。
  - 不含上海履约成本字段（主底稿与桌面表本就无）。

列集（32 列，按业务顺序）：
  商品编号/名称/一二三级类目/合作平台/同步C网零售价/商品产地/供应商/采购员/发货类型/
  各仓库存(北/广/上/武)/平台售价/市场指导价/建议零售价/销项税/
  成本价/毛利率(%)/二次加价率/二次定价(元)/二次定价毛利额(元)/二次定价毛利率(%)/测算采用成本(元)/
  二次定价较平台售价差(元)/成本取值口径/来福池成本(元)/定价状态/被剔除(原)/标注原因(原)

用法：python3 generate_95012_pricing.py
"""
import json, openpyxl

REPO = "/Users/duxun/ai-workspace/projects/benlink-product"
DESK = "/Users/duxun/Desktop/禄品的价格_筛选_2026-11-09.xlsx"
BASE = f"{REPO}/02_报价表/来福商品池_成本及二次定价_2026-10-30.xlsx"
PCODE = f"{REPO}/05_数据与脚本/95012平台核查/核准_禄品4940_P码清单_2026-11-12.json"
OUT = f"{REPO}/02_报价表/禄品供应商模式_95012_成本及二次定价_2026-11-12.xlsx"

# ---- 桌面禄品表（表头在 rows[1]，数据 rows[2:]）----
wb = openpyxl.load_workbook(DESK, read_only=True, data_only=True)
ws = wb.active
drows = list(ws.iter_rows(values_only=True))
dhdr = list(drows[1])
di = {h: i for i, h in enumerate(dhdr)}
# 保留的「原」口径列（排除统一实验列）
DESK_KEEP = [
    "合作平台", "商品编号", "商品名称", "一级类目", "二级类目", "三级类目",
    "同步C网零售价", "商品产地", "供应商", "采购员", "发货类型",
    "北京库存", "广州库存", "上海库存", "武汉库存", "平台售价",
    "成本价", "毛利率(%)", "市场指导价", "建议零售价", "销项税",
    "二次加价率", "二次定价(元)", "二次定价毛利额(元)", "二次定价毛利率(%)",
    "测算采用成本(元)", "二次定价较平台售价差(元)", "被剔除(原)", "标注原因(原)",
]
# 重合时由主底稿覆盖的桌面列
OVERWRITE_FROM_BASE = {
    "测算采用成本(元)": "采用成本(元)",
    "二次定价(元)": "二次定价(元)",
    "二次加价率": "加价率",
}
# 主底稿独有、需新增的补充列
BASE_EXTRA = ["成本取值口径", "来福池成本(元)", "定价状态"]

# 桌面表 dict: pcode -> row
desk_map = {}
for r in drows[2:]:
    p = str(r[di["商品编号"]]).strip()
    if not p:
        continue
    desk_map[p] = r

# ---- 主底稿（表头在 rows[0]，数据 rows[1:]）----
wb2 = openpyxl.load_workbook(BASE, read_only=True, data_only=True)
ws2 = wb2.active
brows = list(ws2.iter_rows(values_only=True))
bhdr = list(brows[0])
bi = {h: i for i, h in enumerate(bhdr)}
base_map = {}
for r in brows[1:]:
    p = str(r[bi["商品编号"]]).strip()
    if not p:
        continue
    base_map[p] = r

# 主底稿权威源列
B_ADOPT = bi["采用成本(元)"]
B_2ND = bi["二次定价(元)"]
B_RATE = bi["加价率"]
B_CALIBER = bi["成本取值口径"]
B_LAIFU = bi["来福池成本(元)"]
B_STATUS = bi["定价状态"]


def valid(v):
    try:
        return v is not None and float(v) > 0
    except (TypeError, ValueError):
        return False


pcodes = json.load(open(PCODE))["pcodes"]
assert len(pcodes) == 4940, f"P码清单应为4940，实际{len(pcodes)}"

# ---- 输出列顺序 ----
OUT_COLS = DESK_KEEP + BASE_EXTRA

wb_out = openpyxl.Workbook()
ws_out = wb_out.active
ws_out.title = "95012禄品供应商模式"
ws_out.append(OUT_COLS)

n_total = 0
n_overlap = 0
n_missing = 0
n_overwrite = 0
for p in pcodes:
    if p not in desk_map:
        raise SystemExit(f"致命：P码 {p} 不在桌面禄品表（骨架缺失）")
    drow = list(desk_map[p])
    base = base_map.get(p)
    if base is not None:
        n_overlap += 1
        # 覆盖采用成本/二次定价/二次加价率（仅当主底稿值有效）
        for desk_col, base_col in OVERWRITE_FROM_BASE.items():
            bv = base[bi[base_col]]
            if valid(bv):
                drow[di[desk_col]] = bv
                n_overwrite += 1
        caliber = base[B_CALIBER]
        laifu = base[B_LAIFU]
        status = base[B_STATUS]
    else:
        n_missing += 1
        caliber = "主底稿无此商品"
        laifu = ""
        status = ""
    out_row = [drow[di[c]] for c in DESK_KEEP] + [caliber, laifu, status]
    ws_out.append(out_row)
    n_total += 1

wb_out.save(OUT)
print(f"输出行数(不含表头): {n_total}")
print(f"  重合(主底稿命中, 权威覆盖): {n_overlap}")
print(f"  缺失(沿用桌面原口径): {n_missing}")
print(f"  权威值覆盖次数(采用成本/二次定价/加价率有效时): {n_overwrite}")
print(f"  缺失行新增列标注为「主底稿无此商品」: {n_missing}")
print(f"已写: {OUT}")
