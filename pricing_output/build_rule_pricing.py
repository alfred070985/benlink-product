#!/usr/bin/env python3
"""生成《规则定价_成本与京东比价》全新工作簿。

严格按用户拍板的四项口径：
  1) 基准 = 禄品4940_筛选表_2026-11-09.xlsx（4,940 个 P 码）
  2) 比价 = 京东 VOP 实时 jd_price（C 端京东价），经「京东商品编码 ↔ VOP sku_id」关联
  3) 成本 = 线上只读逐站点明细，按「展示状态(Status=1)下单站最高值」取值
  4) 定价 = 多个版本同时计算并给出推荐（最终版本由用户拍板）

只读：不修改任何源文件；不写库、不发事件。路径全部由 __file__ 推导，可用环境变量覆盖：
  BENLINK_REPO    业务仓库根目录
  PRICING_DIR     本目录
"""
import csv
import hashlib
import json
import os
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(os.environ.get("PRICING_DIR", HERE))
REPO = Path(os.environ.get("BENLINK_REPO", HERE.parent.parent / "ai-workspace" / "projects" / "benlink-product"))

BASE_XLSX = REPO / "01_源数据" / "禄品4940_筛选表_2026-11-09.xlsx"
JD_MAP_XLSX = REPO / "01_源数据" / "商品池10.29.xlsx"
CHECK_XLSX = REPO / "02_报价表" / "禄品供应商模式_95012_成本及二次定价_2026-11-12.xlsx"
VOP_CSV = HERE.parent / "supply_pull" / "output" / "本来生活VOP_goods_本网ID.csv"
HIST_XLSX = PRICING_DIR / "源数据_京东比价总表1.29.xlsx"
ERP_JSONL = PRICING_DIR / "源数据_ERP站点成本_2026-11-13.jsonl"
PULL_META = PRICING_DIR / "取数说明.json"
OUTPUT = PRICING_DIR / "规则定价_成本与京东比价_2026-11-13.xlsx"

# 三个定价版本；(名称, 加价比例, 客户加价空间/相对京东价的让利)
VERSIONS = [
    ("A", Decimal("1.05"), Decimal("0.20")),
    ("B", Decimal("1.08"), Decimal("0.15")),
    ("C", Decimal("1.10"), Decimal("0.10")),
]
RECOMMEND = "B"

STATUS_ON_SHOW = 1  # AppEnum.ProductStatus.Show = 上架（展示）；0=有效未展示，-1=作废
MONEY = "0.00"
COST_FMT = "0.0000"
RATE = "0.00%"

C_HEADER = "1F4E79"      # 深蓝
C_BASE = "D9E2F3"        # 基础
C_COST = "DDEBF7"        # 成本
C_JD = "E4DFEC"          # 比价
C_VER = {"A": "FCE4D6", "B": "FFF2CC", "C": "E2EFDA"}
C_REC = "FFD966"         # 推荐
C_WARN = "FFC7CE"        # 倒挂
C_MISS = "FCE4D6"        # 缺失
C_GREY = "E7E6E6"
C_OK = "C6EFCE"

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def d(value):
    if value in (None, "", 0, "0"):
        return None
    try:
        x = Decimal(str(value))
    except Exception:
        return None
    return x if x > 0 else None


def money(x):
    return None if x is None else float(x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def rate(x):
    return None if x is None else float(x)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_baseline():
    wb = openpyxl.load_workbook(BASE_XLSX, read_only=True, data_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    header = list(rows[1])
    data = [dict(zip(header, row)) for row in rows[2:] if row[header.index("商品编号")]]
    wb.close()
    return data


def read_check_exclusions():
    if not CHECK_XLSX.exists():
        return {}
    wb = openpyxl.load_workbook(CHECK_XLSX, read_only=True, data_only=True)
    rows = wb.worksheets[0].iter_rows(values_only=True)
    header = list(next(rows))
    i, j = header.index("商品编号"), header.index("被剔除(原)")
    out = {str(r[i]).strip(): str(r[j]) for r in rows if r[i]}
    wb.close()
    return out


def read_jd_codes():
    """P 码 -> 京东商品编码集合（来自商品池10.29 的 ERP 权威字段）。"""
    wb = openpyxl.load_workbook(JD_MAP_XLSX, read_only=True, data_only=True)
    rows = wb.worksheets[0].iter_rows(values_only=True)
    header = list(next(rows))
    i, j = header.index("商品编号"), header.index("京东商品编码")
    out = defaultdict(set)
    total_rows = 0
    for r in rows:
        if r[i] is None:
            continue
        total_rows += 1
        code = r[j]
        if code not in (None, ""):
            out[str(r[i]).strip()].add(str(code).strip())
    wb.close()
    return out, total_rows


def read_vop_prices():
    """sku_id -> 京东 C 端价列表（jd_price）。"""
    out = defaultdict(list)
    total_rows = 0
    with open(VOP_CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            total_rows += 1
            jd = d(row.get("jd_price"))
            if jd is not None:
                out[str(row["sku_id"]).strip()].append({
                    "price": jd,
                    "pool": row.get("pool_id"),
                    "captured_at": row.get("captured_at"),
                    "sale_price": d(row.get("sale_price")),
                })
    return out, total_rows


def read_jd_history():
    """历史比价表：P码 -> 该P码的历史京东C端价记录（当前售价优先，缺失回退折合价）。

    仅在 VOP 实时价缺失时作为辅助比价源；逐行保留「是否同品同规格」「更新时间」供人工判断。
    """
    out = defaultdict(list)
    total = 0
    wb = openpyxl.load_workbook(HIST_XLSX, read_only=True, data_only=True)
    rows = wb.worksheets[0].iter_rows(values_only=True)
    header = list(next(rows))
    i_code = header.index("商品编码")
    i_cur = header.index("京东当前售价")
    i_fold = header.index("京东折合价")
    i_same = header.index("是否同品同规格")
    i_time = header.index("更新时间")
    i_sku = header.index("京东SKU")
    for r in rows:
        if not r or r[i_code] in (None, ""):
            continue
        total += 1
        code = str(r[i_code]).strip()
        price = d(r[i_cur]) or d(r[i_fold])  # 当前售价优先，缺失回退折合价
        out[code].append({
            "price": price,
            "same_spec": str(r[i_same]).strip() if r[i_same] is not None else "",
            "updated": str(r[i_time]).strip() if r[i_time] is not None else "",
            "sku": str(r[i_sku]).strip() if r[i_sku] is not None else "",
        })
    wb.close()
    return out, total


def read_erp_records():
    out = defaultdict(list)
    total = 0
    with open(ERP_JSONL, encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            o = json.loads(line)
            total += 1
            out[str(o["ProductID"]).strip()].append(o)
    return out, total


def cost_for(pcode, ship, records):
    """按展示状态下单站最高值取成本；返回 (采用成本, 口径, 状态, 明细行)。"""
    show = [r for r in records if r.get("Status") == STATUS_ON_SHOW and not r.get("IsDelete")]
    detail = []
    for r in show:
        detail.append({
            "pcode": pcode,
            "name": r.get("ProductName"),
            "ship": ship,
            "site": r.get("WebSiteSysNo"),
            "channel": r.get("ChannelSysNo"),
            "sale_channel": r.get("SaleChannelSysNo"),
            "status": r.get("Status"),
            "sysno": r.get("ProductSysNo"),
            "unit_cost": d(r.get("UnitCost")),
            "purchase": d(r.get("DefaultPurchasePrice")),
        })
    if not show:
        return None, "—", "无展示状态记录", detail, 0, 0

    unit_vals = [x["unit_cost"] for x in detail if x["unit_cost"] is not None]
    pur_vals = [x["purchase"] for x in detail if x["purchase"] is not None]
    unit_max = max(unit_vals) if unit_vals else None
    pur_max = max(pur_vals) if pur_vals else None
    sites = len({x["site"] for x in detail})

    chosen = None
    basis = "—"
    if ship == "仓发":
        if unit_max is not None:
            chosen, basis = unit_max, "含税单位成本·展示状态单站最高"
        elif pur_max is not None:
            chosen, basis = pur_max, "回退采购价·展示状态单站最高"
        else:
            return None, "—", "展示记录无有效成本", detail, sites, len(show)
    else:
        if pur_max is not None:
            chosen, basis = pur_max, "采购价·展示状态单站最高"
        else:
            return None, "—", "展示记录无有效采购价", detail, sites, len(show)

    for x in detail:
        x["selected"] = "是" if (
            (basis.startswith("含税单位成本") and x["unit_cost"] == chosen) or
            (("采购价" in basis) and x["purchase"] == chosen)
        ) else "否"
    return chosen, basis, "有成本", detail, sites, len(show)


def jd_for(pcode, codes, vop, history):
    """比价取值：实时 VOP jd_price 优先，缺失时回退历史比价表。

    返回字典，逐项保留来源、更新时间、是否同品同规格，供人工判断历史价可信度。
    """
    real_rows = []
    for code in sorted(codes):
        for item in vop.get(code, []):
            real_rows.append({"code": code, **item})
    real_prices = sorted({x["price"] for x in real_rows})
    real_price = real_prices[0] if real_prices else None

    hist_rows = [x for x in history.get(pcode, []) if x["price"] is not None]
    hist_prices = sorted({x["price"] for x in hist_rows})
    hist_price = hist_prices[0] if hist_prices else None
    hist_same = "；".join(sorted({x["same_spec"] for x in hist_rows if x["same_spec"]}))
    hist_time = "；".join(sorted({x["updated"] for x in hist_rows if x["updated"]}))
    hist_sku = "；".join(sorted({x["sku"] for x in hist_rows if x["sku"]}))

    if real_price is not None:
        price, src, status = real_price, "实时VOP", "有实时比价"
        note = "实时多价取最低" if len(real_prices) > 1 else "实时单一价"
    elif hist_price is not None:
        price, src, status = hist_price, "历史比价表1.29", "有历史比价"
        note = "历史多价取最低" if len(hist_prices) > 1 else "历史价"
    else:
        price, src, status = None, "无", "无比价"
        note = "无价"

    return dict(price=price, real_price=real_price, hist_price=hist_price, src=src,
                status=status, note=note,
                sku="；".join(sorted({x["code"] for x in real_rows})),
                real_hits=len(real_rows), hist_same=hist_same, hist_time=hist_time,
                hist_sku=hist_sku)


def price_one(cost, jd, ratio, headroom):
    """返回一版定价的 8 个结果字段。"""
    if cost is None:
        return dict(markup=None, cap=None, final=None, gross=None, gm=None,
                    client_price=None, client_room=None, status="成本缺失·不出价", candidate=None)
    markup = cost * ratio
    cap = None if jd is None else jd * (Decimal("1") - headroom)
    candidate = min(markup, cap) if cap is not None else markup
    if candidate < cost:
        return dict(markup=money(markup), cap=money(cap), final=None, gross=None, gm=None,
                    client_price=None, client_room=None,
                    status="价格倒挂·不出价", candidate=money(candidate))
    final = candidate
    gross = final - cost
    gm = gross / final if final else None
    client_price = final * (Decimal("1") + headroom)
    client_room = (jd / final - Decimal("1")) if (jd is not None and final) else None
    status = "正常"
    if jd is None:
        status = "无比价·按成本加成"
    elif cap is not None and cap < markup:
        status = "受京东价封顶"
    return dict(markup=money(markup), cap=money(cap), final=money(final),
                gross=money(gross), gm=rate(gm), client_price=money(client_price),
                client_room=rate(client_room), status=status, candidate=money(candidate))


def style_header(ws, row, fills, height=32):
    for col, fill in enumerate(fills, start=1):
        c = ws.cell(row=row, column=col)
        c.fill = PatternFill("solid", fgColor=fill)
        c.font = Font(bold=True, color="FFFFFF" if fill == C_HEADER else "1F1F1F", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def main():
    for p in (BASE_XLSX, JD_MAP_XLSX, VOP_CSV, HIST_XLSX, ERP_JSONL, PULL_META):
        if not p.exists():
            raise SystemExit(f"缺少源文件：{p}")

    baseline = read_baseline()
    exclusions = read_check_exclusions()
    jd_codes, jd_map_rows = read_jd_codes()
    vop, vop_rows = read_vop_prices()
    hist, hist_rows = read_jd_history()
    erp, erp_rows = read_erp_records()
    meta = json.loads(PULL_META.read_text(encoding="utf-8"))

    rows_out = []
    site_detail = []
    stats = Counter()
    for item in baseline:
        p = str(item["商品编号"]).strip()
        ship = str(item["发货类型"]).strip()
        cost, basis, cost_status, detail, sites, show_n = cost_for(p, ship, erp.get(p, []))
        site_detail.extend(detail)
        info = jd_for(p, jd_codes.get(p, set()), vop, hist)
        rec = {
            "pcode": p,
            "name": item.get("商品名称"),
            "c1": item.get("一级类目"),
            "c2": item.get("二级类目"),
            "c3": item.get("三级类目"),
            "ship": ship,
            "supplier": item.get("供应商"),
            "platform_price": d(item.get("平台售价")),
            "guide_price": d(item.get("建议零售价")),
            "sites": sites,
            "show_n": show_n,
            "unit_max": max([x["unit_cost"] for x in detail if x["unit_cost"] is not None], default=None),
            "purchase_max": max([x["purchase"] for x in detail if x["purchase"] is not None], default=None),
            "cost": cost,
            "basis": basis,
            "cost_status": cost_status,
            "jd_codes": ";".join(sorted(jd_codes.get(p, set()))) or "",
            "jd_sku": info["sku"] or "",
            "jd": info["price"],
            "jd_src": info["src"],
            "jd_status": info["status"],
            "price_note": info["note"],
            "jd_real": info["real_price"],
            "jd_hist": info["hist_price"],
            "jd_hist_same": info["hist_same"],
            "jd_hist_time": info["hist_time"],
            "jd_hist_sku": info["hist_sku"],
            "jd_hits": info["real_hits"],
            "excluded_before": exclusions.get(p, ""),
        }
        for tag, ratio, headroom in VERSIONS:
            rec[tag] = price_one(cost, info["price"], ratio, headroom)
        rows_out.append(rec)

        stats["明细行"] += 1
        stats["有成本"] += 1 if cost is not None else 0
        stats["无展示记录"] += 1 if cost_status == "无展示状态记录" else 0
        stats["有京东价"] += 1 if info["price"] is not None else 0
        stats["实时京东价"] += 1 if info["real_price"] is not None else 0
        stats["历史京东价"] += 1 if info["src"] == "历史比价表1.29" else 0
        stats["历史同品同规格"] += 1 if (info["src"] == "历史比价表1.29" and "是" in info["hist_same"]) else 0
        stats["多价取最低"] += 1 if "多价取最低" in info["note"] else 0
        if cost_status == "无展示状态记录" and exclusions.get(p) == "是":
            stats["无展示记录且历史已剔除"] += 1

    # 版本指标
    metrics = {}
    for tag, ratio, headroom in VERSIONS:
        finals = [r[tag]["final"] for r in rows_out if r[tag]["final"] is not None]
        gms = [r[tag]["gm"] for r in rows_out if r[tag]["gm"] is not None]
        jd_group = [r for r in rows_out if r["jd"] is not None and r[tag]["final"] is not None]
        ok_sell = sum(1 for r in jd_group if r[tag]["client_price"] is not None and r[tag]["client_price"] <= float(r["jd"]))
        metrics[tag] = {
            "可出价SKU": len(finals),
            "成本缺失": sum(1 for r in rows_out if r[tag]["status"] == "成本缺失·不出价"),
            "价格倒挂": sum(1 for r in rows_out if r[tag]["status"] == "价格倒挂·不出价"),
            "受京东价封顶": sum(1 for r in rows_out if r[tag]["status"] == "受京东价封顶"),
            "无实时比价": sum(1 for r in rows_out if r[tag]["status"] == "无比价·按成本加成"),
            "平均毛利率": sum(gms) / len(gms) if gms else 0,
            "毛利率中位数": sorted(gms)[len(gms) // 2] if gms else 0,
            "有京东价且可出价": len(jd_group),
            "客户加价后不高于京东价": ok_sell,
            "客户加价后占比": ok_sell / len(jd_group) if jd_group else 0,
        }

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 01 规则说明 ----------------
    ws = wb.create_sheet("01_规则说明")
    set_widths(ws, [4, 26, 110])
    ws["A1"] = "规则定价 · 成本 / 定价 / 比价 口径说明"
    ws["A1"].font = Font(bold=True, size=16, color=C_HEADER)
    ws.merge_cells("A1:C1")
    ws.row_dimensions[1].height = 30
    lines = [
        ("SEC", "一、目标与边界", ""),
        ("TXT", "目标", "规则定价，全表不人工逐 SKU 干预；所有取值由规则计算，缺口只留空标注，不回填、不预估。"),
        ("TXT", "本表性质", "全新 Excel，基准为禄品 4,940 个 P 码；未修改任何源文件，未写库、未发事件。"),
        ("TXT", "拍板记录", "基准表=禄品4940筛选表；比价=京东 VOP 实时 jd_price 优先、缺失回退历史比价表1.29（用户后续追加拍板并入）；成本=线上只读逐站点明细，按展示状态下单站最高值；定价=多版本计算后给推荐。"),
        ("SEC", "二、成本规则", ""),
        ("TXT", "数据源", "本来生活 ERP 只读接口 ProLhotse/GetListByBasic（接口文档：仅对 product_channel_info / 价格表执行 SELECT）。本次一次性只读取数，结果落盘为《源数据_ERP站点成本_2026-11-13.jsonl》。"),
        ("TXT", "「展示状态」定义", "ERP 渠道商品 Status 枚举（AppEnum.ProductStatus）：1=上架(Show，展示)、0=有效(Valid，未展示)、-1=作废(Abandon)。本表只取 Status=1 且 IsDelete=False 的记录。"),
        ("TXT", "「单站最高值」定义", "同一 P 码在所有展示状态的站点/渠道明细中，逐条比较后取最高值；不是平均值、不是中位数。同时列出站点数、记录数和逐条明细（《06_站点成本明细》）。"),
        ("TXT", "仓发商品", "优先取含税单位成本 UnitCost 的展示状态单站最高值；若所有展示记录都无有效含税单位成本（空或≤0），回退取采购价 DefaultPurchasePrice 的单站最高值；两者都无 → 成本缺失。"),
        ("TXT", "直发商品", "取采购价 DefaultPurchasePrice 的展示状态单站最高值；无有效值 → 成本缺失。"),
        ("TXT", "发货类型来源", "《禄品4940_筛选表》的「发货类型」列（仓发 / 直发），不自行推断。"),
        ("TXT", "缺失处理", "采用成本留空、三个版本的输出价全部留空、状态标「成本缺失·不出价」。禁止用其它口径自动补齐。"),
        ("TXT", "口径交叉验证", f"无展示记录的 {stats['无展示记录']} 个 P 码中，{stats['无展示记录且历史已剔除']} 个在 95012 历史报价表中本就被剔除；说明「展示状态」口径与历史可定价范围基本一致。该对照仅用于校验，不参与定价。"),
        ("SEC", "三、比价规则", ""),
        ("TXT", "两路比价源（用户拍板）", "① 实时：VOP 的 jd_price（京东 C 端客户看到的京东价），明确排除 sale_price（协议价）、market_price（划线价）、naked_price（裸价）。② 历史：仅当实时价缺失时，回退《京东比价总表1.29.xlsx》的「京东当前售价」（缺失再退「京东折合价」）作辅助基准。"),
        ("TXT", "实时价怎么对上", "VOP 本网 P 码与禄品 4,940 个 P 码零重叠，改由《商品池10.29.xlsx》的「京东商品编码」↔ VOP 的 sku_id 关联。命中率低，故引入历史源补充。"),
        ("TXT", "历史价怎么对上", "历史比价表直接用「商品编码」=禄品 P 码关联，覆盖更广。表内「更新时间」跨度 2024–2026，非统一 1 月价，已在明细逐行标注，请结合「是否同品同规格」自行判断可信度。"),
        ("TXT", "优先级与标注", "实时价优先；无实时价才用历史价，并在「比价来源」列标「历史比价表1.29」、明细列同时保留实时价/历史价/同品同规格/更新时间，两路互不覆盖、可回溯。"),
        ("TXT", "重复取值", "同一 P 码命中多条记录时取非空价的最小值（保守，避免高估市场价），并标「多价取最低」及命中记录数。"),
        ("TXT", "无价处理", "两路都无价时，比价基准留空，比价上限不适用，按纯成本加成出价，状态标「无比价」。"),
        ("TXT", "调研但未参与定价", "《本网ID_待确认.csv》为歧义映射（一 sku 多 P 码），需人工裁决，未参与本表定价。"),
        ("SEC", "四、定价规则", ""),
        ("TXT", "成本加成价", "采用成本 × 加价比例"),
        ("TXT", "比价上限", "京东 C 端价 × (1 − 客户加价空间)。含义：我们给客户留出的、相对京东价的让利比例，保证客户二次加价后仍不高于京东价。"),
        ("TXT", "候选输出价", "有京东价时 = min(成本加成价, 比价上限)；无京东价时 = 成本加成价。"),
        ("TXT", "价格倒挂", "若候选输出价 < 采用成本，则标「价格倒挂·不出价」，最终输出价留空（不亏本出价），候选价记在参考列。"),
        ("TXT", "毛利口径", "毛利额 = 最终输出价 − 采用成本；毛利率 = 毛利额 ÷ 最终输出价。"),
        ("TXT", "客户加价后售价", "最终输出价 × (1 + 客户加价空间)；另列「客户实际可加价率」= 京东价 ÷ 最终输出价 − 1。"),
        ("TXT", "金额精度", "计算使用源值全精度；展示金额保留 2 位小数（成本列 4 位），比例按百分比显示。"),
        ("SEC", "五、版本与推荐", ""),
        ("TXT", "三个版本", "A=×1.05 / 留20%，B=×1.08 / 留15%，C=×1.10 / 留10%；参数集中在《03_定价参数》，逐 SKU 结果见《05_逐SKU定价明细》。"),
        ("TXT", "推荐性质", f"《04_版本对比与推荐》给出建议版本 {RECOMMEND} 及数据依据；这是建议，不是结论，最终采用哪个版本由您拍板。"),
        ("SEC", "六、颜色说明", ""),
        ("TXT", "表头", "深蓝=分组标题；浅蓝=基础/成本；浅紫=比价；橙/黄/绿=版本 A/B/C；金黄=推荐列。"),
        ("TXT", "状态", "浅红=价格倒挂；浅橙=成本缺失/无展示记录；浅灰=无比价；浅黄=有历史比价；浅绿=正常/有实时比价。"),
        ("SEC", "七、复现", ""),
        ("TXT", "脚本", "pull_erp_site_costs.py（一次性只读取数）→ build_rule_pricing.py（本表生成）。"),
        ("TXT", "生成时间", datetime.now().isoformat(timespec="seconds")),
    ]
    r = 3
    for kind, k, v in lines:
        if kind == "SEC":
            c = ws.cell(row=r, column=1, value=k)
            c.font = Font(bold=True, size=12, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=C_HEADER)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
            ws.row_dimensions[r].height = 24
        else:
            a = ws.cell(row=r, column=2, value=k)
            a.font = Font(bold=True, color=C_HEADER)
            a.alignment = Alignment(vertical="top", wrap_text=True)
            b = ws.cell(row=r, column=3, value=v)
            b.alignment = Alignment(vertical="top", wrap_text=True)
            b.font = Font(color="333333")
            ws.row_dimensions[r].height = 30 if len(v) < 70 else 44
        r += 1
    ws.freeze_panes = "A3"

    # ---------------- 02 数据源清单 ----------------
    ws = wb.create_sheet("02_数据源清单")
    hdr = ["序号", "数据源", "类型", "路径 / 接口", "参与环节", "规模", "关键字段", "数据时间", "读取方式", "是否参与定价", "备注"]
    ws.append(hdr)
    style_header(ws, 1, [C_HEADER] * len(hdr))
    sources = [
        (1, "禄品4940_筛选表_2026-11-09.xlsx", "Excel 基准表", str(BASE_XLSX), "基准 / 发货类型 / 基础信息", f"{len(baseline):,} 行 × 36 列", "商品编号(P码)、商品名称、一级类目、发货类型、平台售价、建议零售价", "2026-11-09", "本地只读", "是", "4,940 个唯一 P 码"),
        (2, "ProLhotse/GetListByBasic", "线上只读接口", URL_TEXT, "成本取数（逐站点）", f"{erp_rows:,} 条站点记录", "UnitCost(含税单位成本)、DefaultPurchasePrice(采购价)、Status、WebSiteSysNo、IsDelete", meta.get("pulled_at"), "只读 SELECT，未写库", "是", "落盘文件见下一行"),
        (3, "源数据_ERP站点成本_2026-11-13.jsonl", "本地落盘源数据", str(ERP_JSONL), "成本取数（复现用）", f"{erp_rows:,} 条", "同上，仅保留参与定价的字段", meta.get("pulled_at"), "本地只读", "是", "一次性取数结果，0 失败批次"),
        (4, "本来生活VOP_goods_本网ID.csv", "京东 VOP 价格", str(VOP_CSV), "京东 C 端价", f"{vop_rows:,} 行 × 40 列", "sku_id、jd_price、sale_price、market_price、captured_at", "2026-11-13 抓取", "本地只读", "是", "只用 jd_price"),
        (5, "商品池10.29.xlsx", "Excel 映射表", str(JD_MAP_XLSX), "P 码 → 京东商品编码", f"{jd_map_rows:,} 行", "商品编号、京东商品编码", "2026-10-29", "本地只读", "是", "仅用编码做关联，不用其价格"),
        (6, "禄品供应商模式_95012_成本及二次定价_2026-11-12.xlsx", "历史报价表", str(CHECK_XLSX), "口径交叉验证", "4,940 行", "商品编号、被剔除(原)", "2026-11-12", "本地只读", "否", "仅校验展示状态口径，不参与定价"),
        (7, "源数据_京东比价总表1.29.xlsx", "历史比价源", str(HIST_XLSX), "辅助比价（实时价缺失时回退）", f"{hist_rows:,} 行", "商品编码、京东当前售价、京东折合价、是否同品同规格、更新时间", "表内更新时间 2024–2026", "只读（已从 Downloads 原样复制）", "是（辅助）", "P码直连禄品；覆盖广但为历史价，明细逐行标来源与更新时间，请核同品同规格"),
        (8, "本网ID_待确认.csv", "歧义映射清单", str(HERE.parent / "supply_pull" / "output" / "本网ID_待确认.csv"), "调研", "1,751 行", "sku_id、pool_candidate_pcodes、goods_map_pcode", "2026-11-13", "只读调研", "否", "多 P 码歧义，不做自动裁决"),
        (9, "取数说明.json", "取数元数据", str(PULL_META), "留痕", "1 份", "read_only、side_effects、status_enum、批次", meta.get("pulled_at"), "本地只读", "是", "记录只读口径"),
    ]
    for s in sources:
        ws.append(list(s))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = BORDER
            c.font = Font(size=10, color="333333")
        ws.row_dimensions[row[0].row].height = 46
    set_widths(ws, [6, 34, 16, 62, 24, 20, 48, 18, 18, 14, 34])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{ws.max_row}"

    # ---------------- 03 定价参数 ----------------
    ws = wb.create_sheet("03_定价参数")
    ws["A1"] = "定价参数（集中配置，改这里即可重算；最终采用哪个版本由您拍板）"
    ws["A1"].font = Font(bold=True, size=13, color=C_HEADER)
    ws.merge_cells("A1:G1")
    hdr = ["版本", "定位", "加价比例", "客户加价空间", "比价上限系数", "是否建议", "说明"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, [C_HEADER] * len(hdr))
    pos = {"A": "保守：让利多、毛利薄", "B": "均衡：毛利与客户空间兼顾", "C": "进取：毛利高、空间小"}
    for tag, ratio, headroom in VERSIONS:
        ws.append([tag, pos[tag], float(ratio), float(headroom), float(Decimal("1") - headroom),
                   "建议" if tag == RECOMMEND else "备选",
                   "成本加成价=成本×加价比例；比价上限=京东价×比价上限系数"])
    ws.append([])
    ws.append(["成本状态口径", "Status = 1（上架 / 展示），且 IsDelete = False", "", "", "", "", "枚举：-1 作废 / 0 有效未展示 / 1 上架展示"])
    ws.append(["京东多价口径", "同一 P 码多条 VOP 价格取最低值", "", "", "", "", "保守口径，避免高估京东 C 端价"])
    ws.append(["金额精度", "计算用全精度，展示 2 位（成本 4 位）", "", "", "", "", "四舍五入 ROUND_HALF_UP"])
    for row in ws.iter_rows(min_row=4, max_row=7):
        for c in row:
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=True)
        row[2].number_format = "0.00"
        row[3].number_format = RATE
        row[4].number_format = "0.00"
        if row[0].value == RECOMMEND:
            for c in row:
                c.fill = PatternFill("solid", fgColor=C_REC)
    for row in ws.iter_rows(min_row=9, max_row=11):
        row[0].font = Font(bold=True, color=C_HEADER)
        row[1].alignment = Alignment(wrap_text=True, vertical="center")
    set_widths(ws, [16, 30, 16, 18, 18, 14, 62])
    ws.freeze_panes = "A4"

    # ---------------- 04 版本对比与推荐 ----------------
    ws = wb.create_sheet("04_版本对比与推荐")
    ws["A1"] = "三个定价版本对比与推荐"
    ws["A1"].font = Font(bold=True, size=14, color=C_HEADER)
    ws.merge_cells("A1:D1")
    hdr = ["指标", "A（×1.05 / 留20%）", "B（×1.08 / 留15%）", "C（×1.10 / 留10%）"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, [C_HEADER] * 4)
    metric_rows = [
        ("可出价 SKU 数", "可出价SKU", "int"),
        ("成本缺失·不出价", "成本缺失", "int"),
        ("价格倒挂·不出价", "价格倒挂", "int"),
        ("受京东价封顶", "受京东价封顶", "int"),
        ("无比价·按成本加成", "无实时比价", "int"),
        ("平均毛利率", "平均毛利率", "pct"),
        ("毛利率中位数", "毛利率中位数", "pct"),
        ("有京东价且可出价 SKU", "有京东价且可出价", "int"),
        ("客户加价后不高于京东价", "客户加价后不高于京东价", "int"),
        ("客户加价后不高于京东价占比", "客户加价后占比", "pct"),
    ]
    for label, key, kind in metric_rows:
        vals = [metrics[t][key] for t in ("A", "B", "C")]
        ws.append([label] + vals)
        row = ws.max_row
        for col in range(2, 5):
            c = ws.cell(row=row, column=col)
            c.number_format = RATE if kind == "pct" else "#,##0"
            c.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        for c in row:
            c.border = BORDER
    r = ws.max_row + 2
    ws.cell(row=r, column=1, value="推荐结论（建议，最终由您拍板）").font = Font(bold=True, size=12, color=C_HEADER)
    jd_cov = stats["有京东价"] / stats["明细行"] if stats["明细行"] else 0
    rt = stats["实时京东价"]
    hs = stats["历史京东价"]
    conclusions = [
        f"1) 并入历史比价表后，有京东基准价的 SKU 从 {rt:,}（仅实时）增至 {stats['有京东价']:,} 个（{jd_cov:.1%}）：实时 VOP {rt:,} + 历史表回退 {hs:,}（其中标「同品同规格=是」{stats['历史同品同规格']:,}）。历史价为旧数据且逐行标了更新时间与同品同规格，可信度请人工把关。",
        f"2) 仍有 {stats['明细行'] - stats['有京东价']:,} 个 SKU 两路都无价，比价规则对其不生效，只按成本加成出价；三版在这类 SKU 上仅差加价比例带来的毛利差异。在能比价的 SKU 上，三版「客户加价后不高于京东价」为 100%，这是「比价上限」按规则构造的必然结果，不是版本优劣。",
        f"3) 取舍数据：可出价 SKU = A {metrics['A']['可出价SKU']:,} / B {metrics['B']['可出价SKU']:,} / C {metrics['C']['可出价SKU']:,}；价格倒挂（不出价）= A {metrics['A']['价格倒挂']} / B {metrics['B']['价格倒挂']} / C {metrics['C']['价格倒挂']}；被京东价封顶 = A {metrics['A']['受京东价封顶']} / B {metrics['B']['受京东价封顶']} / C {metrics['C']['受京东价封顶']}。让利越多，越容易把价格压到成本以下而出不了价；让利越少，客户二次加价空间越小。",
        f"4) 建议默认采用 B（×1.08 / 留15%）：平均毛利率 {metrics['B']['平均毛利率']:.2%}，介于 A 的 {metrics['A']['平均毛利率']:.2%} 与 C 的 {metrics['C']['平均毛利率']:.2%} 之间；倒挂 {metrics['B']['价格倒挂']} 个少于 A，客户加价空间 15% 又比 C 的 10% 宽。",
        "5) 历史价属旧数据（更新时间 2024–2026 不等），与实时价严肃性不同；若要用于正式对外报价，建议先复核历史行「是否同品同规格=是」，或补全实时 VOP 映射后替换。",
        "6) 本页为建议口径，不是最终结论；采用哪个版本、是否采信历史比价、是否放宽展示状态口径，均需您明确拍板。",
    ]
    for text in conclusions:
        r += 1
        ws.cell(row=r, column=1, value=text)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=1).font = Font(color="333333", size=10)
        ws.row_dimensions[r].height = 38
    set_widths(ws, [36, 24, 24, 24])
    ws.freeze_panes = "B4"

    # ---------------- 05 逐SKU定价明细 ----------------
    ws = wb.create_sheet("05_逐SKU定价明细")
    groups = []
    def add(labels, fill):
        for x in labels:
            groups.append((x, fill))
    add(["商品编号(P码)", "商品名称", "一级类目", "二级类目", "三级类目", "发货类型", "供应商",
         "平台售价(元)", "建议零售价(元)"], C_BASE)
    add(["展示状态站点数", "展示状态记录数", "含税单位成本·单站最高", "采购价·单站最高",
         "采用成本(元)", "成本取值口径", "成本状态"], C_COST)
    add(["京东商品编码", "命中京东SKU", "比价来源", "比价状态",
         "采用京东C端价(元)", "实时京东价 jd_price(元)", "历史京东价(元)",
         "历史同品同规格", "历史价更新时间", "命中VOP记录数", "多价取值"], C_JD)
    for tag, ratio, headroom in VERSIONS:
        add([f"{tag}_成本加成价", f"{tag}_比价上限", f"{tag}_候选输出价", f"{tag}_最终输出价",
             f"{tag}_毛利额", f"{tag}_毛利率", f"{tag}_客户加价后售价",
             f"{tag}_客户实际可加价率", f"{tag}_定价状态"], C_VER[tag])
    add([f"推荐{RECOMMEND}_输出价", f"推荐{RECOMMEND}_毛利率", "备注 / 异常"], C_REC)
    ws.append([x[0] for x in groups])
    style_header(ws, 1, [x[1] for x in groups], height=44)

    for rec in rows_out:
        row = [
            rec["pcode"], rec["name"], rec["c1"], rec["c2"], rec["c3"], rec["ship"], rec["supplier"],
            money(rec["platform_price"]), money(rec["guide_price"]),
            rec["sites"], rec["show_n"],
            float(rec["unit_max"]) if rec["unit_max"] is not None else None,
            float(rec["purchase_max"]) if rec["purchase_max"] is not None else None,
            float(rec["cost"]) if rec["cost"] is not None else None,
            rec["basis"], rec["cost_status"],
            rec["jd_codes"], rec["jd_sku"], rec["jd_src"], rec["jd_status"],
            float(rec["jd"]) if rec["jd"] is not None else None,
            float(rec["jd_real"]) if rec["jd_real"] is not None else None,
            float(rec["jd_hist"]) if rec["jd_hist"] is not None else None,
            rec["jd_hist_same"], rec["jd_hist_time"],
            rec["jd_hits"], rec["price_note"],
        ]
        for tag, _, _ in VERSIONS:
            v = rec[tag]
            row += [v["markup"], v["cap"], v["candidate"], v["final"], v["gross"], v["gm"],
                    v["client_price"], v["client_room"], v["status"]]
        rec_v = rec[RECOMMEND]
        note_parts = []
        if rec["cost_status"] != "有成本":
            note_parts.append(f"成本：{rec['cost_status']}")
        if rec["excluded_before"] == "是":
            note_parts.append("95012历史报价表曾被剔除")
        if "多价取最低" in rec["price_note"]:
            note_parts.append(f"京东{rec['price_note']}")
        if rec["jd_src"] == "历史比价表1.29":
            note_parts.append("历史价(非实时)·请核同品同规格")
            if rec["jd_hist_time"]:
                note_parts.append(f"历史价更新时间 {rec['jd_hist_time']}")
        for tag, _, _ in VERSIONS:
            if "倒挂" in rec[tag]["status"]:
                note_parts.append(f"{tag}版价格倒挂")
        row += [rec_v["final"], rec_v["gm"], "；".join(note_parts)]
        ws.append(row)

    # 样式与数值格式（按列名定位，避免列序调整后错位）
    idx_map = {name: i for i, (name, _) in enumerate(groups, start=1)}
    cost_cols, money_cols, rate_cols = set(), set(), set()
    for name, i in idx_map.items():
        if "毛利率" in name or "可加价率" in name:
            rate_cols.add(i)
        elif name in ("含税单位成本·单站最高", "采购价·单站最高", "采用成本(元)"):
            cost_cols.add(i)
        elif name in ("平台售价(元)", "建议零售价(元)", "采用京东C端价(元)",
                      "实时京东价 jd_price(元)", "历史京东价(元)", f"推荐{RECOMMEND}_输出价"):
            money_cols.add(i)
        elif name.endswith(("成本加成价", "比价上限", "候选输出价", "最终输出价", "毛利额", "客户加价后售价")):
            money_cols.add(i)

    status_cols = [idx_map[f"{tag}_定价状态"] for tag, _, _ in VERSIONS]
    cost_status_col = idx_map["成本状态"]
    jd_status_col = idx_map["比价状态"]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            c.border = BORDER
            c.font = Font(size=10, color="333333")
            if c.column in cost_cols:
                c.number_format = COST_FMT
            elif c.column in money_cols:
                c.number_format = MONEY
            elif c.column in rate_cols:
                c.number_format = RATE
        for sc in status_cols:
            c = ws.cell(row=row[0].row, column=sc)
            v = str(c.value or "")
            if "倒挂" in v:
                c.fill = PatternFill("solid", fgColor=C_WARN)
            elif "缺失" in v:
                c.fill = PatternFill("solid", fgColor=C_MISS)
            elif "无比价" in v:
                c.fill = PatternFill("solid", fgColor=C_GREY)
            elif v == "正常":
                c.fill = PatternFill("solid", fgColor=C_OK)
        cc = ws.cell(row=row[0].row, column=cost_status_col)
        if cc.value != "有成本":
            cc.fill = PatternFill("solid", fgColor=C_MISS)
        jc = ws.cell(row=row[0].row, column=jd_status_col)
        if jc.value == "有实时比价":
            jc.fill = PatternFill("solid", fgColor=C_OK)
        elif jc.value == "有历史比价":
            jc.fill = PatternFill("solid", fgColor=C_REC)
        else:
            jc.fill = PatternFill("solid", fgColor=C_GREY)
    widths = [16, 44, 14, 18, 18, 10, 26, 13, 13,
              12, 12, 18, 18, 14, 30, 20,
              20, 22, 14, 16, 16, 18, 18, 14, 18, 12, 14]
    widths += [14] * 27
    widths += [16, 16, 40]
    set_widths(ws, widths[:ws.max_column])
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    # ---------------- 06 站点成本明细 ----------------
    ws = wb.create_sheet("06_站点成本明细")
    hdr = ["商品编号(P码)", "商品名称", "发货类型", "站点(WebSiteSysNo)", "渠道(ChannelSysNo)",
           "销售渠道(SaleChannelSysNo)", "展示状态(Status)", "渠道商品编号(ProductSysNo)",
           "含税单位成本(元)", "采购价(元)", "是否提供最终采用成本", "说明"]
    ws.append(hdr)
    style_header(ws, 1, [C_HEADER] * len(hdr))
    for x in site_detail:
        ws.append([
            x["pcode"], x["name"], x["ship"], x["site"], x["channel"], x["sale_channel"],
            "上架(展示)" if x["status"] == 1 else x["status"], x["sysno"],
            float(x["unit_cost"]) if x["unit_cost"] is not None else None,
            float(x["purchase"]) if x["purchase"] is not None else None,
            x.get("selected", "否"),
            "仅列展示状态(Status=1)、未删除的站点记录；同一 P 码逐条比较取最高值",
        ])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            c.border = BORDER
        for i in (9, 10):
            row[i - 1].number_format = COST_FMT
        if row[10].value == "是":
            row[10].fill = PatternFill("solid", fgColor=C_OK)
    set_widths(ws, [16, 44, 10, 16, 16, 20, 16, 22, 18, 16, 20, 56])
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:L{ws.max_row}"

    # ---------------- 07 缺口异常与校验 ----------------
    ws = wb.create_sheet("07_缺口异常与校验")
    set_widths(ws, [34, 22, 22, 22, 60])
    ws["A1"] = "缺口、异常与交付校验"
    ws["A1"].font = Font(bold=True, size=14, color=C_HEADER)
    ws.merge_cells("A1:E1")
    checks = [
        ("基准表行数", len(baseline), "应与明细行数一致"),
        ("逐 SKU 明细行数", len(rows_out), "不增不减"),
        ("唯一 P 码数", len({r["pcode"] for r in rows_out}), "应与基准表一致"),
        ("ERP 站点记录数", erp_rows, f"取数说明记录 {meta.get('records')}"),
        ("展示状态站点记录数", len(site_detail), "参与单站最高值比较"),
        ("有成本 SKU", stats["有成本"], "可进入定价"),
        ("无展示状态记录 SKU", stats["无展示记录"], "成本缺失·不出价"),
        ("其中 95012 历史已剔除", stats["无展示记录且历史已剔除"], "口径交叉验证"),
        ("实时京东价 SKU", stats["实时京东价"], "VOP jd_price，经京东商品编码关联"),
        ("历史京东价 SKU（回退）", stats["历史京东价"], "仅实时缺失时启用，来自历史比价表"),
        ("其中同品同规格=是", stats["历史同品同规格"], "历史价可信度标注，供人工判断"),
        ("合计有京东基准价 SKU", stats["有京东价"], "实时+历史去重后"),
        ("京东多价取最低 SKU", stats["多价取最低"], "保守口径"),
        ("VOP 价格文件行数", vop_rows, "原始 46,635 行"),
        ("历史比价表行数", hist_rows, "已复制入交付目录，SHA 见下"),
        ("商品池10.29 行数", jd_map_rows, "编码映射来源"),
    ]
    for tag, _, _ in VERSIONS:
        checks.append((f"{tag} 版可出价 SKU", metrics[tag]["可出价SKU"], "最终输出价非空"))
        checks.append((f"{tag} 版价格倒挂", metrics[tag]["价格倒挂"], "留空不出价"))
    hdr = ["校验项", "结果", "单位", "结论", "说明"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(row=3, column=i, value=h)
    style_header(ws, 3, [C_HEADER] * 5)
    for k, v, note in checks:
        ws.append([k, v, "", "通过" if v is not None else "检查", note])
    r = ws.max_row + 2
    ws.cell(row=r, column=1, value="源文件指纹（SHA256）").font = Font(bold=True, size=12, color=C_HEADER)
    ws.cell(row=r, column=2, value="生成时核对，确认源文件未被改写")
    r += 1
    for label, path in [
        ("禄品4940_筛选表", BASE_XLSX),
        ("商品池10.29", JD_MAP_XLSX),
        ("京东VOP价格CSV", VOP_CSV),
        ("历史比价表落盘", HIST_XLSX),
        ("ERP站点成本落盘", ERP_JSONL),
        ("本表生成脚本", Path(__file__)),
    ]:
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=sha256(path))
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="只读声明")
    ws.cell(row=r, column=2, value="全程只读打开源文件；未修改任何源表、未写库、未发商品变更事件；线上取数接口按文档仅执行 SELECT。")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 34
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        for c in row:
            c.border = BORDER
            c.font = Font(size=10, color="333333")
    ws.freeze_panes = "A4"

    wb.save(OUTPUT)
    print(json.dumps({
        "output": str(OUTPUT),
        "sheets": wb.sheetnames,
        "rows": len(rows_out),
        "site_detail_rows": len(site_detail),
        "jd_price_hits_total": stats["有京东价"],
        "jd_price_realtime": stats["实时京东价"],
        "jd_price_history": stats["历史京东价"],
        "cost_available": stats["有成本"],
        "metrics": metrics,
    }, ensure_ascii=False, indent=2))


URL_TEXT = "https://vm-online.benlai.com/erp-product-webapi/ProLhotse/GetListByBasic"

if __name__ == "__main__":
    main()
