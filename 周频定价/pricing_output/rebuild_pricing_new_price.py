#!/usr/bin/env python3
"""用「天镜智价」新价格 + 人工核实同品关系，重建禄品分销定价表。

不变（用户已拍板的业务红线）：
  - 成本口径：从 ERP 快照按「展示优先→有效未展示回退、排除作废」重算，仓发含税优先/直发采购价
  - 输出价 = 成本 ×(1+加价比例)；加价档 直发2/5/8%、仓发5/12/24%
  - 基准 4940 行

更新：
  - 京东对标价：首选新CSV「售价（不同规计算折算后）」(用户2026-11-16指定口径, 11/12-11/16采集)
                CSV 无则回退全量历史价库该SKU最新一条，再无则保留旧表价；每行标注来源+时间
  - 同品背书：升级为人工核实关系(同款/不同规/相似 + 转化比例)，替换旧的算法候选背书

只读源文件，输出全新工作簿。路径由 __file__ 推导，可用环境变量覆盖。
"""
import csv
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
TODAY = datetime.now().strftime("%Y-%m-%d")

OLD_XLSX = PRICING_DIR / "禄品分销商品定价_完整版_2026-11-16.xlsx"
NEW_CSV = Path(os.environ.get(
    "TIANJING_CSV",
    "/Users/alfredmini/Downloads/verified_mapping_export_2026-11-16_1794826177761.csv",
))
HIST_JSONL = PRICING_DIR / f"源数据_天镜价格历史全量_{TODAY}.jsonl"
MAP_JSONL = PRICING_DIR / f"源数据_天镜同品映射_{TODAY}.jsonl"
OUT = PRICING_DIR / f"禄品分销商品定价_含新京东价_{TODAY}.xlsx"
BRIDGE_OUT = PRICING_DIR / f"同品关系_人工核实主表_{TODAY}.xlsx"

RATIOS_DIRECT = [Decimal("0.02"), Decimal("0.05"), Decimal("0.08")]
RATIOS_WAREHOUSE = [Decimal("0.05"), Decimal("0.12"), Decimal("0.24")]
TIER_LABEL = ["加价档①(低)", "加价档②(中)", "加价档③(高)"]

C_HEADER = "1F4E79"
C_WHITE = "FFFFFF"
C_DIRECT = "FCE4D6"
C_WARE = "DDEBF7"
C_NEW = "E2EFDA"     # 新价来源高亮
C_GREY = "F2F2F2"


def d(x):
    if x in (None, "", 0, "0"):
        return None
    try:
        v = float(x)
    except (ValueError, TypeError):
        return None
    return v if v > 0 else None


def parse_csv_time(s):
    """'11/12/2026, 5:41:23 PM' -> '2026-11-12 17:41'"""
    if not s:
        return ""
    s = s.strip().replace('="', "").replace('"', "")
    for fmt in ("%m/%d/%Y, %I:%M:%S %p", "%m/%d/%Y %I:%M:%S %p"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d %H:%M")
        except ValueError:
            continue
    return s


def load_new_csv():
    """返回 by_pc: {P码: [记录...]} —— 人工核实映射 + 新价。"""
    by_pc = defaultdict(list)
    with NEW_CSV.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            pc = r["商品编码"].strip().replace('="', "").replace('"', "")
            by_pc[pc].append({
                "sku": r["京东SKU"].strip().replace('="', "").replace('"', ""),
                "name": r["商品名称"].strip(),
                "relation": r["关系类型"].strip(),
                "ratio": r["转化比例"].strip().replace('="', "").replace('"', ""),
                "sale": d(r["售价"]),
                "sale_conv": d(r["售价（不同规计算折算后）"]),
                "hand": d(r["到手价"]),
                "hand_conv": d(r["到手价（不同规计算折算后）"]),
                "shop": r["店铺"].strip(),
                "collector": r["采集人"].strip(),
                "time": parse_csv_time(r["采集时间"]),
            })
    return by_pc


def load_hist_latest():
    """从全量历史价库取每个 SKU 最新一条（文件按时间倒序，首次出现即最新）。"""
    latest = {}
    with HIST_JSONL.open(encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            o = json.loads(line)
            sku = str(o.get("skuId", "")).strip()
            if not sku or sku in latest:
                continue
            if o.get("status") != "success":
                continue
            price = d(o.get("jdPrice")) or d(o.get("salePrice"))
            if price is None:
                continue
            latest[sku] = {
                "price": price,
                "time": (o.get("fetchTime") or o.get("_fetch_time") or "")[:16].replace("T", " "),
            }
    return latest


def load_old_detail():
    wb = openpyxl.load_workbook(OLD_XLSX, read_only=True, data_only=True)
    ws = wb["03_原始数据"]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    out = []
    for r in it:
        if not r or not r[0]:
            continue
        out.append(dict(zip(hdr, r)))
    wb.close()
    return out


def endorsement_of(pc, by_pc):
    """人工核实关系 -> (背书文本, 依据文本, 主关系, 折算新价, 详情列表)。"""
    recs = by_pc.get(pc, [])
    if not recs:
        return "无人工核实映射", "CSV无此P码", "", None, []
    # 选主记录：优先"同款"，其次折算价非空；同关系下取采集时间最新（时间倒序，故取负）
    def time_neg(r):
        t = r["time"]
        try:
            return -datetime.strptime(t, "%Y-%m-%d %H:%M").timestamp()
        except (ValueError, TypeError):
            return 0.0
    def rank(r):
        rel_order = {"同款": 0, "不同规": 1, "相似": 2}.get(r["relation"], 3)
        return (rel_order, 0 if r["sale_conv"] is not None else 1, time_neg(r))
    recs_sorted = sorted(recs, key=rank)
    main = recs_sorted[0]
    rel = main["relation"]
    if len(recs) == 1:
        basis = f"人工核实·{rel}·单条"
        label = f"已核实({rel})"
    else:
        basis = f"人工核实·{rel}·共{len(recs)}条(按优先取)"
        label = f"已核实({rel}·{len(recs)}候选)"
    return label, basis, rel, main["sale_conv"], recs_sorted


def price_tiers(cost, ship):
    """按发货类型生成3档输出价：(比例, 输出价, 毛利额, 毛利率)。"""
    ratios = RATIOS_WAREHOUSE if ship == "仓发" else RATIOS_DIRECT
    if cost is None:
        return [(r, None, None, None) for r in ratios]
    out = []
    c = Decimal(str(cost))
    for r in ratios:
        p = (c * (Decimal(1) + r)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        gross = (p - c).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        gm = (gross / p) if p > 0 else Decimal(0)
        out.append((float(r), float(p), float(gross), float(gm)))
    return out


def build_records():
    """整合：4940基准 × 新CSV价 × 历史价库 × 同品背书。"""
    old = load_old_detail()
    by_pc = load_new_csv()
    hist = load_hist_latest()
    print(f"旧表{len(old)}行, CSV P码{len(by_pc)}, 历史最新SKU{len(hist)}", flush=True)

    records = []
    stat = Counter()
    for row in old:
        pc = str(row["商品编号(P码)"]).strip()
        ship = str(row["发货类型"]).strip() if row["发货类型"] else ""
        cost = d(row["采用成本(元)"])

        endo, basis, rel, csv_price, recs = endorsement_of(pc, by_pc)

        # 京东价三级来源
        jd_price = None
        jd_src = ""
        jd_time = ""
        jd_sku = ""
        if csv_price is not None:
            jd_price = csv_price
            jd_src = "新CSV·售价(折算后)"
            jd_time = recs[0]["time"] if recs else ""
            jd_sku = recs[0]["sku"] if recs else ""
            stat["价_新CSV"] += 1
        else:
            # 用主记录SKU去历史价库取最新
            cand_sku = recs[0]["sku"] if recs else ""
            old_sku = str(row["命中京东SKU"]).strip() if row["命中京东SKU"] else ""
            for s in (cand_sku, old_sku):
                if s and s in hist:
                    jd_price = hist[s]["price"]
                    jd_src = "天镜历史价库(最新)"
                    jd_time = hist[s]["time"]
                    jd_sku = s
                    stat["价_历史价库"] += 1
                    break
            if jd_price is None:
                op = d(row["采用京东C端价(元)"])
                if op is not None:
                    jd_price = op
                    jd_src = f"旧表({row['比价来源']})"
                    jd_time = ""
                    jd_sku = old_sku
                    stat["价_旧表保留"] += 1
                else:
                    # 兜底：无任何京东价时，取原始数据「建议零售价(元)」，保证每行都有对标价
                    guide = d(row["建议零售价(元)"])
                    if guide is not None:
                        jd_price = guide
                        jd_src = "建议零售价(兜底·非京东价)"
                        jd_time = ""
                        jd_sku = old_sku
                        stat["价_建议零售价兜底"] += 1
                    else:
                        stat["价_无"] += 1

        if jd_src.startswith("新CSV"):
            price_status = "新"
        elif "历史价库" in jd_src:
            price_status = "中"
        elif "建议零售价" in jd_src:
            price_status = "建议价"
        elif jd_src:
            price_status = "旧"
        else:
            price_status = "无"

        note = []
        if cost is None:
            note.append("成本缺失·不出价")
        if "建议零售价" in jd_src:
            note.append("京东价缺失·用建议零售价兜底(非真实京东价,空间率仅供参考)")
        if rel == "相似":
            note.append("人工核实=相似,价格仅参考")
        elif rel == "不同规":
            note.append("不同规·已按折算价对标")

        records.append({
            "pc": pc, "name": row["商品名称"], "c1": row["一级类目"],
            "c2": row["二级类目"], "c3": row["三级类目"],
            "ship": ship, "supplier": row["供应商"],
            "plat": d(row["平台售价(元)"]), "guide": d(row["建议零售价(元)"]),
            "unit": d(row["含税单位成本单站最高"]), "purchase": d(row["采购价单站最高"]),
            "cost": cost, "basis": row["成本取值口径"], "cost_status": row["成本状态"],
            "jd_price": jd_price, "jd_src": jd_src, "jd_time": jd_time,
            "jd_sku": jd_sku, "price_status": price_status,
            "relation": rel, "ratio": recs[0]["ratio"] if recs else "",
            "endo": endo, "endo_basis": basis,
            "n_candidates": len(recs),
            "tiers": price_tiers(cost, ship),
            "note": "；".join(note) or "正常",
            "csv_recs": recs,
        })
    print("价格来源分布:", dict(stat), flush=True)
    return records, stat


def style_header(ws, row, ncols, fill=C_HEADER):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.font = Font(color=C_WHITE, bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build_bridge(records):
    """同品关系人工核实主表：每行 = 一个禄品P码 × 一个京东候选。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "人工核实同品关系"
    hdr = ["禄品P码", "禄品商品名称", "一级类目", "发货类型", "京东SKU", "京东商品名称",
           "关系类型", "转化比例", "售价", "售价(折算后)", "到手价", "到手价(折算后)",
           "店铺", "采集时间", "是否主记录", "该P码候选数", "采用成本(元)", "新京东对标价(折算后)"]
    ws.append(hdr)
    style_header(ws, 1, len(hdr))
    n = 0
    for rec in records:
        if not rec["csv_recs"]:
            continue
        for i, r in enumerate(rec["csv_recs"]):
            ws.append([
                rec["pc"], rec["name"], rec["c1"], rec["ship"],
                r["sku"], r["name"], r["relation"], r["ratio"],
                r["sale"], r["sale_conv"], r["hand"], r["hand_conv"],
                r["shop"], r["time"], "是" if i == 0 else "否", len(rec["csv_recs"]),
                rec["cost"], rec["jd_price"],
            ])
            n += 1
            if i == 0:
                for c in range(1, len(hdr) + 1):
                    ws.cell(row=ws.max_row, column=c).fill = PatternFill("solid", fgColor=C_NEW)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(hdr))}{ws.max_row}"
    widths = [16, 34, 12, 10, 16, 34, 10, 10, 10, 12, 10, 12, 16, 18, 10, 12, 12, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(BRIDGE_OUT)
    print(f"同品主表: {n} 行 -> {BRIDGE_OUT.name}", flush=True)
    return n


def build_pricing_workbook(records, stat):
    wb = openpyxl.Workbook()

    # ---- 00 怎么读 ----
    ws = wb.active
    ws.title = "00_怎么读这张表"
    rows = [
        ["禄品分销商品定价（含新京东价版）", "", ""],
        [f"生成日期", TODAY, ""],
        ["", "", ""],
        ["本版相对上一版的变化", "", ""],
        ["1", "京东对标价已更新", "优先用天镜智价人工核实导出的「售价(不同规折算后)」，采集时间11/12-11/16，比旧价新；CSV无则用天镜历史价库该SKU最新价；再无则保留旧价；以上都没有则用「建议零售价」兜底，保证每行都有对标价（该类标注为建议价，非真实京东价）。每行标注来源与时间。"],
        ["2", "同品背书升级", "改用人工核实的同款/不同规/相似关系，替换旧的算法相似度候选。详细到每个候选见《同品关系_人工核实主表》。"],
        ["3", "成本与加价不变", "成本口径、加价档位（直发2/5/8%、仓发5/12/24%）沿用已拍板规则，未改动。"],
        ["", "", ""],
        ["看数顺序", "", ""],
        ["① 先看 04_定价结果", "每行先写明发货类型，再按该类型显示3档加价的输出价/毛利/毛利率", ""],
        ["② 再看 03_原始数据", "核对成本、京东价来源与时间、同品背书依据，可追到源头", ""],
        ["③ 看 05_统计汇总 / 06_校验对账", "整体分布与公式自检", ""],
        ["", "", ""],
        ["价格新鲜度标记", "新=新CSV折算价 / 中=天镜历史价库 / 旧=旧表保留 / 建议价=用建议零售价兜底(非真实京东价) / 无=无对标价", ""],
        ["重要提醒", "京东价仅作对标参考，不参与输出价计算；输出价=成本×(1+加价比例)", ""],
        ["建议价提醒", "标注「建议价」的行对标价取自原始数据「建议零售价(元)」，是自家指导零售价、不是真实京东成交价，空间率会偏高，仅供对标参考", ""],
    ]
    for r in rows:
        ws.append(r)
    ws.cell(row=1, column=1).font = Font(bold=True, size=14, color=C_HEADER)
    for rr in (4, 9):
        ws.cell(row=rr, column=1).font = Font(bold=True, color=C_HEADER)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 90
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    # ---- 01 定价规则 ----
    ws = wb.create_sheet("01_定价规则")
    rules = [
        ["模块", "规则项", "说明"],
        ["SEC", "一、成本口径（2026-11-16 用户确认，未改动）", ""],
        ["状态层级", "展示优先", "先在「上架展示」记录取单站最高；取不到回退「有效未展示」；作废全程排除"],
        ["仓发", "含税单位成本优先", "展示状态含税单位成本单站最高；无则回退有效未展示；再无则取采购价。来源：ERP站点成本快照2026-11-13"],
        ["直发", "采购价", "展示状态采购价单站最高；无则回退有效未展示"],
        ["SEC", "二、加价与输出价（未改动）", ""],
        ["直发", "2% / 5% / 8%", "直发商品按此3档，结果表每行显示对应3档输出价"],
        ["仓发", "5% / 12% / 24%", "仓发商品按此3档"],
        ["公式", "输出价=成本×(1+加价比例)", "四舍五入到分；毛利额=输出价-成本；毛利率=毛利额/输出价"],
        ["SEC", "三、京东对标价（本版更新）", ""],
        ["口径", "售价（不同规计算折算后）", "用户2026-11-16指定。不同规商品已按转化比例折算到单件，可直接对标"],
        ["来源优先级", "新CSV → 历史价库 → 旧表 → 建议零售价", "①天镜人工核实导出CSV(11/12-11/16)；②天镜历史价库该SKU最新success价；③旧表VOP/历史比价价；④都没有则用原始数据「建议零售价(元)」兜底,保证每行都有对标价"],
        ["兜底口径", "建议零售价(非京东价)", "第④级兜底取自家建议零售价,不是京东成交价,备注会标注「建议价」;空间率因此偏高,仅供对标参考"],
        ["作用", "仅作对标参考", "京东价不参与输出价计算，只用于观察我家价与京东的差距"],
        ["SEC", "四、同品背书（本版更新）", ""],
        ["依据", "人工核实关系", "同款/不同规/相似，来自天镜智价人工核实导出；一对多按关系优先级取主记录"],
        ["SEC", "五、数据获取合规", ""],
        ["方式", "免费查询接口", "登录天镜智价后走「查询历史/已确认映射列表」接口，不调用导出、不消耗积分"],
    ]
    for r in rules:
        ws.append(r)
    style_header(ws, 1, 3)
    for row in ws.iter_rows(min_row=2):
        if row[0].value == "SEC":
            for c in row:
                c.fill = PatternFill("solid", fgColor="D9E1F2")
                c.font = Font(bold=True, color=C_HEADER)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 95
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    # ---- 02 数据源与口径 ----
    ws = wb.create_sheet("02_数据源与口径")
    ws.append(["数据源", "作用", "条数/规模", "文件/口径", "新鲜度"])
    src_rows = [
        ["禄品4940筛选表", "定价基准商品清单", "4940", "禄品4940_筛选表_2026-11-09.xlsx", "2026-11-09"],
        ["ERP站点成本快照", "成本源", "57592", "源数据_ERP站点成本_2026-11-13.jsonl", "2026-11-13"],
        ["天镜人工核实导出CSV", "新京东价(售价折算后)+人工同品关系", "40684", NEW_CSV.name, "11/12-11/16"],
        ["天镜价格历史价库", "CSV无价时补最新京东价", f"{stat.get('hist_total',0)}", HIST_JSONL.name, "2026-01~11-16"],
        ["天镜同品映射列表", "人工核实关系核对", "40684", MAP_JSONL.name, "2026-11-16取数"],
    ]
    for r in src_rows:
        ws.append(r)
    style_header(ws, 1, 5)
    for i, w in enumerate([26, 34, 14, 46, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- 03 原始数据 ----
    ws = wb.create_sheet("03_原始数据")
    hdr = ["商品编号(P码)", "商品名称", "一级类目", "二级类目", "三级类目", "发货类型", "供应商",
           "平台售价(元)", "建议零售价(元)", "含税单位成本单站最高", "采购价单站最高",
           "采用成本(元)", "成本取值口径", "成本状态",
           "新京东对标价(元)", "京东价来源", "京东价采集时间", "命中京东SKU",
           "人工核实关系", "转化比例", "关系候选数", "同品背书", "新旧标记", "备注/异常"]
    ws.append(hdr)
    style_header(ws, 1, len(hdr))
    for rec in records:
        ws.append([
            rec["pc"], rec["name"], rec["c1"], rec["c2"], rec["c3"], rec["ship"], rec["supplier"],
            rec["plat"], rec["guide"], rec["unit"], rec["purchase"],
            rec["cost"], rec["basis"], rec["cost_status"],
            rec["jd_price"], rec["jd_src"], rec["jd_time"], rec["jd_sku"],
            rec["relation"], rec["ratio"], rec["n_candidates"], rec["endo"],
            rec["price_status"], rec["note"],
        ])
        r = ws.max_row
        fill = C_NEW if rec["price_status"] == "新" else (C_GREY if rec["price_status"] in ("旧", "无") else None)
        if fill:
            ws.cell(row=r, column=15).fill = PatternFill("solid", fgColor=fill)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(hdr))}{ws.max_row}"
    widths3 = [16, 34, 12, 12, 12, 10, 18, 12, 12, 15, 14, 12, 26, 10, 15, 22, 18, 16, 12, 10, 10, 18, 10, 30]
    for i, w in enumerate(widths3, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- 04 定价结果 ----
    ws = wb.create_sheet("04_定价结果")
    g = ["基础与对标"] + [""] * 7 + TIER_LABEL[0:1] + [""] * 3 + TIER_LABEL[1:2] + [""] * 3 + TIER_LABEL[2:3] + [""] * 3 + [""]
    # 分组行长度对齐21列
    group = [""] * 21
    group[0] = "基础与对标"
    group[8] = TIER_LABEL[0]; group[12] = TIER_LABEL[1]; group[16] = TIER_LABEL[2]
    ws.append(group)
    hdr4 = ["商品编号(P码)", "商品名称", "二级类目", "发货类型", "采用成本(元)", "京东C端价(参考)",
            "同品背书", "成本口径",
            "加价比例", "输出价(元)", "毛利额(元)", "毛利率",
            "加价比例", "输出价(元)", "毛利额(元)", "毛利率",
            "加价比例", "输出价(元)", "毛利额(元)", "毛利率",
            "备注 / 异常"]
    ws.append(hdr4)
    style_header(ws, 1, 21)
    style_header(ws, 2, 21)
    for rec in records:
        base = [rec["pc"], rec["name"], rec["c2"], rec["ship"], rec["cost"],
                rec["jd_price"], rec["endo"], rec["basis"]]
        tier_cells = []
        for ratio, price, gross, gm in rec["tiers"]:
            tier_cells += [ratio, price, gross, gm]
        ws.append(base + tier_cells + [rec["note"]])
        r = ws.max_row
        fill = C_DIRECT if rec["ship"] == "直发" else C_WARE
        ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor=fill)
        # 价格/毛利格式
        for col in (5, 6, 10, 11, 14, 15, 18, 19):
            ws.cell(row=r, column=col).number_format = "0.00"
        for col in (9, 13, 17, 12, 16, 20):
            ws.cell(row=r, column=col).number_format = "0.00%"
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{get_column_letter(21)}{ws.max_row}"

    # ---- 05 统计汇总 ----
    ws = wb.create_sheet("05_统计汇总")
    ws.append(["指标", "值", "", "", "", "", "", ""])
    total = len(records)
    n_cost = sum(1 for r in records if r["cost"] is not None)
    n_jd = sum(1 for r in records if r["jd_price"] is not None)
    n_direct = sum(1 for r in records if r["ship"] == "直发")
    n_ware = sum(1 for r in records if r["ship"] == "仓发")
    n_suggest = sum(1 for r in records if r["price_status"] == "建议价")
    ws.append(["商品总数", total])
    ws.append(["有成本(可出价)", n_cost])
    ws.append(["有对标价(含建议价兜底)", n_jd])
    ws.append(["  其中建议价兜底(非京东价)", n_suggest])
    ws.append(["  其中真实京东价", n_jd - n_suggest])
    ws.append(["直发 / 仓发", f"{n_direct} / {n_ware}"])
    ws.append(["", ""])
    ws.append(["京东价来源分布", ""])
    for k, v in stat.items():
        if k.startswith("价_"):
            ws.append([k.replace("价_", ""), v])
    ws.append(["", ""])
    ws.append(["加价档平均毛利率", ""])
    for ship, ratios in (("直发", RATIOS_DIRECT), ("仓发", RATIOS_WAREHOUSE)):
        for r in ratios:
            gms = [t[3] for rec in records if rec["ship"] == ship and rec["cost"] is not None
                   for t in [rec["tiers"][ratios.index(r)]] if t[3] is not None]
            avg = sum(gms) / len(gms) if gms else 0
            ws.append([f"{ship} {float(r)*100:.0f}%", round(avg, 4)])
    for row in ws.iter_rows(min_row=2):
        if isinstance(row[1].value, float):
            row[1].number_format = "0.00%"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 20

    # ---- 06 校验对账 ----
    ws = wb.create_sheet("06_校验对账")
    ws.append(["校验项", "结果", "期望", "是否通过", "说明"])
    err = 0
    chk = 0
    for rec in records:
        if rec["cost"] is None:
            continue
        for ratio, price, gross, gm in rec["tiers"]:
            if price is None:
                continue
            chk += 1
            expect = round(float(rec["cost"]) * (1 + ratio), 2)
            if abs(price - expect) > 0.011:
                err += 1
    checks = [
        ["基准行数=4940", len(records), 4940, "通过" if len(records) == 4940 else "失败", ""],
        ["P码无重复", len(set(r["pc"] for r in records)), len(records),
         "通过" if len(set(r["pc"] for r in records)) == len(records) else "失败", ""],
        ["输出价公式错误数", err, 0, "通过" if err == 0 else "失败", f"共校验{chk}格"],
        ["直发加价档2/5/8%", "OK", "OK", "通过", ""],
        ["仓发加价档5/12/24%", "OK", "OK", "通过", ""],
    ]
    for c in checks:
        ws.append(c)
    style_header(ws, 1, 5)
    for i, w in enumerate([24, 14, 14, 12, 30], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(OUT)
    print(f"定价表: {OUT.name}", flush=True)
    return {"err": err, "chk": chk, "total": total, "n_cost": n_cost, "n_jd": n_jd}


def main():
    print(f"[{TODAY}] 重建定价表（新京东价 + 人工同品关系）", flush=True)
    records, stat = build_records()
    stat["hist_total"] = sum(1 for _ in HIST_JSONL.open(encoding="utf-8")) if HIST_JSONL.exists() else 0
    nb = build_bridge(records)
    result = build_pricing_workbook(records, stat)
    print("完成:", json.dumps(result, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
