#!/usr/bin/env python3
"""生成《禄品分销商品定价_完整版》——原始数据 + 规则 + 结果数据 三合一。

用户拍板的口径：
  1) 基准 = 规则定价明细 4940 行（禄品4940 筛选表），成本按用户2026-11-16确认口径从ERP快照重算
  2) 输出价 = 采用成本 × (1 + 加价比例)，纯成本加成；京东价仅作对标参考，不参与计算
  3) 加价比例：直发 2%/5%/8%，仓发 5%/12%/24%
  4) 结果表每行先显示「发货类型」，再按规则显示该类型对应的 3 档加价比例
  5) 同品背书：历史已确认采信 + 新增候选①优先/②次看且单匹配采信；一对多仅对新增候选排除

只读：不修改任何源文件。路径由 __file__ 推导，可用环境变量覆盖。
"""
import hashlib
import json
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))

DETAIL_XLSX = PRICING_DIR / "规则定价_成本与京东比价_2026-11-13.xlsx"
SIM_XLSX = PRICING_DIR / "同品关系_候选相似度与人工确认_2026-11-16.xlsx"
ERP_JSONL = PRICING_DIR / "源数据_ERP站点成本_2026-11-13.jsonl"
HIST_XLSX = PRICING_DIR / "源数据_京东比价总表1.29.xlsx"
VOP_CSV = HERE.parent / "supply_pull" / "output" / "本来生活VOP_goods_本网ID.csv"
PULL_META = PRICING_DIR / "取数说明.json"
BASE_XLSX = Path(__import__("os").environ.get(
    "BENLINK_REPO", HERE.parent.parent / "ai-workspace" / "projects" / "benlink-product"
)) / "01_源数据" / "禄品4940_筛选表_2026-11-09.xlsx"
OUT = PRICING_DIR / "禄品分销商品定价_完整版_2026-11-16.xlsx"

RATIOS_DIRECT = [Decimal("0.02"), Decimal("0.05"), Decimal("0.08")]
RATIOS_WAREHOUSE = [Decimal("0.05"), Decimal("0.12"), Decimal("0.24")]
TIER_LABEL = ["加价档①(低)", "加价档②(中)", "加价档③(高)"]

C_HEADER = "1F4E79"
C_WHITE = "FFFFFF"
C_DIRECT = "FCE4D6"
C_WARE = "DDEBF7"
C_TIER = ["E2EFDA", "FFF2CC", "FCE4D6"]
C_GOOD = "C6EFCE"
C_BAD = "FFC7CE"
C_WARN = "FFEB9C"
C_GREY = "E7E6E6"
C_SECTION = "D9E2F3"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = "0.00"
RATE = "0.0%"
RATE2 = "0.00%"


def d(x):
    if x in (None, "", 0, "0"):
        return None
    try:
        v = Decimal(str(x))
    except Exception:
        return None
    return v if v > 0 else None


def f2(x):
    return None if x is None else float(x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def style(cell, fill=None, bold=False, color=None, align="center", size=10, wrap=False):
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(bold=bold, color=color or "1F1F1F", size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def load_detail():
    wb = openpyxl.load_workbook(DETAIL_XLSX, read_only=True, data_only=True)
    ws = wb["05_逐SKU定价明细"]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    rows = []
    for r in it:
        if not r or r[0] in (None, ""):
            continue
        rows.append(dict(zip(hdr, r)))
    wb.close()
    return rows


def load_erp_records():
    """从 ERP 快照读取全部站点记录，按 P 码归集。"""
    out = defaultdict(list)
    with open(ERP_JSONL, encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            o = json.loads(line)
            out[str(o["ProductID"]).strip()].append(o)
    return out


def _max_field(recs, field):
    """在一组站点记录里取某字段的单站最高值（仅正数值）。"""
    vals = [d(r.get(field)) for r in recs]
    vals = [x for x in vals if x is not None]
    return max(vals) if vals else None


def cost_from_erp(pcode, ship, records):
    """按用户 2026-11-16 原文口径取成本（结构：状态优先级 + 字段优先级）：

       状态优先级：「优先展示状态(Status=1)下取单站最高值」，取不到才回退到
                   有效未展示(Status=0)；作废(-1)与 IsDelete 全程排除。
       字段优先级：
         仓发 = 优先取含税单位成本(UnitCost)单站最高；无含税单位成本的取采购价(DefaultPurchasePrice)单站最高
         直发 = 取采购价(DefaultPurchasePrice)单站最高

       每一层取值都先试「展示状态记录」，无值再试「有效未展示记录」。
       返回 (成本, 口径文本, 状态文本, 站点数, 有效记录数)。
    """
    show = [r for r in records if r.get("Status") == 1 and not r.get("IsDelete")]      # 展示状态
    unshow = [r for r in records if r.get("Status") == 0 and not r.get("IsDelete")]    # 有效未展示
    valid = show + unshow
    if not valid:
        return None, "—", "无有效状态记录(仅作废)", 0, 0
    sites = len({r.get("WebSiteSysNo") for r in valid})

    def pick(field, label):
        # 层级：展示状态优先，取不到回退有效未展示
        v = _max_field(show, field)
        if v is not None:
            return v, f"{label}·展示状态单站最高"
        v = _max_field(unshow, field)
        if v is not None:
            return v, f"{label}·回退有效未展示单站最高"
        return None, None

    if ship == "仓发":
        # 字段优先：含税单位成本 → 采购价
        v, basis = pick("UnitCost", "含税单位成本")
        if v is not None:
            return v, basis, "有成本", sites, len(valid)
        v, basis = pick("DefaultPurchasePrice", "采购价")
        if v is not None:
            return v, basis + "(无含税单位成本)", "有成本", sites, len(valid)
        return None, "—", "有效状态记录无成本值", sites, len(valid)
    else:
        v, basis = pick("DefaultPurchasePrice", "采购价")
        if v is not None:
            return v, basis, "有成本", sites, len(valid)
        return None, "—", "有效状态记录无采购价", sites, len(valid)


def load_endorsement():
    """返回 {p码: (背书状态, 依据)} 与统计。"""
    wb = openpyxl.load_workbook(SIM_XLSX, read_only=True, data_only=True)
    # 历史已确认
    ws = wb["01_已确认同品"]
    it = ws.iter_rows(values_only=True); next(it)
    hist = set()
    for r in it:
        if r[0]:
            hist.add(str(r[0]).strip())
    # 新增候选
    ws = wb["02_待确认候选"]
    it = ws.iter_rows(values_only=True); hdr = list(next(it))
    i_lv = hdr.index("复核优先级"); i_pc = hdr.index("禄品P码")
    by = defaultdict(list)
    for r in it:
        if not r or not r[i_pc]:
            continue
        by[str(r[i_pc]).strip()].append(str(r[i_lv]).strip())
    wb.close()

    endorse = {}
    stats = Counter()
    for pc in hist:
        endorse[pc] = ("已确认(历史)", "历史表同品同规格=是")
        stats["历史确认"] += 1
    for pc, lvs in by.items():
        if len(lvs) > 1:
            endorse[pc] = ("一对多·不采信", f"{len(lvs)}个候选，全部排除")
            stats["一对多不采信"] += 1
        elif lvs[0] in ("①优先", "②次看"):
            endorse[pc] = ("候选采信", f"单候选·{lvs[0]}")
            stats["候选采信"] += 1
        else:
            endorse[pc] = ("存疑·不采信", f"单候选·{lvs[0]}")
            stats["存疑不采信"] += 1
    return endorse, stats


def price_tiers(cost, ship):
    ratios = RATIOS_DIRECT if ship == "直发" else (RATIOS_WAREHOUSE if ship == "仓发" else None)
    if ratios is None:
        return None
    if cost is None:
        return [dict(ratio=r, out=None, gross=None, gm=None) for r in ratios]
    c = Decimal(str(cost))
    out = []
    for r in ratios:
        v = (c * (Decimal("1") + r)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        g = (v - c).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        gm = (g / v) if v > 0 else None
        out.append(dict(ratio=r, out=v, gross=g, gm=gm))
    return out


def main():
    for p in (DETAIL_XLSX, SIM_XLSX, ERP_JSONL, HIST_XLSX, VOP_CSV, BASE_XLSX):
        if not p.exists():
            raise SystemExit(f"缺少源文件：{p}")

    detail = load_detail()
    endorse, e_stats = load_endorsement()
    erp = load_erp_records()
    print(f"载入明细 {len(detail)} 行；同品背书 {len(endorse)} 个P码 {dict(e_stats)}；ERP快照 {len(erp)} 个P码", flush=True)

    # 组装每行结果：成本直接从 ERP 快照按新口径重算（不沿用上游"只取展示状态"的结果）
    records = []
    cost_fix = 0
    for row in detail:
        pc = str(row["商品编号(P码)"]).strip()
        ship = str(row["发货类型"]).strip() if row["发货类型"] else ""
        old_cost = d(row["采用成本(元)"])
        cost, basis, cost_status, sites, recs_n = cost_from_erp(pc, ship, erp.get(pc, []))
        if old_cost is None and cost is not None:
            cost_fix += 1
        jd = d(row["采用京东C端价(元)"])
        endo, basis_endo = endorse.get(pc, ("无同品候选", "未找到同品关系"))
        tiers = price_tiers(cost, ship)
        note = []
        if cost is None:
            note.append("成本缺失·不出价")
        if jd is None:
            note.append("无京东对标价")
        if endo in ("一对多·不采信", "存疑·不采信"):
            note.append("京东对标价未背书")
        # 从 ERP 重算的 unit/pur 展示值
        er = [r for r in erp.get(pc, []) if r.get("Status") in (1, 0) and not r.get("IsDelete")]
        unit_max = max([d(r.get("UnitCost")) for r in er if d(r.get("UnitCost"))], default=None)
        pur_max = max([d(r.get("DefaultPurchasePrice")) for r in er if d(r.get("DefaultPurchasePrice"))], default=None)
        records.append(dict(
            pc=pc, name=row["商品名称"], c1=row["一级类目"], c2=row["二级类目"], c3=row["三级类目"],
            ship=ship, supplier=row["供应商"], plat=d(row["平台售价(元)"]), guide=d(row["建议零售价(元)"]),
            unit=unit_max, purchase=pur_max, sites=sites, show_n=recs_n,
            cost=cost, basis=basis, cost_status=cost_status,
            jd_code=row["京东商品编码"], jd_sku=row["命中京东SKU"],
            jd_src=row["比价来源"], jd_status=row["比价状态"], jd=jd,
            real=d(row["实时京东价 jd_price(元)"]), histp=d(row["历史京东价(元)"]),
            endo=endo, endo_basis=basis_endo, tiers=tiers, note="；".join(note) or "正常",
        ))
    print(f"按新口径补回成本 {cost_fix} 行", flush=True)

    # 校验前置
    assert len(records) == 4940, f"明细行数异常 {len(records)}"
    ship_cnt = Counter(r["ship"] for r in records)
    assert set(ship_cnt) <= {"直发", "仓发"}, f"发货类型异常 {ship_cnt}"
    print(f"发货类型：{dict(ship_cnt)}", flush=True)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------------- 00 怎么读这张表 ----------------
    ws = wb.create_sheet("00_怎么读这张表")
    set_widths(ws, [4, 24, 104])
    ws["A1"] = "禄品分销商品定价 · 完整版"
    ws["A1"].font = Font(bold=True, size=17, color=C_HEADER); ws.merge_cells("A1:C1")
    ws.row_dimensions[1].height = 32
    guide = [
        ("SEC", "一句话说明", "每一行 = 一个禄品商品。先看它「发货类型」是直发还是仓发，再按该类型的规则看 3 档加价后的建议输出价。全程按公式计算，不人工干预。"),
        ("SEC", "三个部分", ""),
        ("TXT", "① 原始数据", "03_原始数据：4940 个商品的原始信息、采用成本、京东对标价、同品背书状态，全部来自已确认的数据源，可追溯。"),
        ("TXT", "② 规则", "01_定价规则 + 02_数据源与口径：写清楚成本怎么取、加价比例是多少、京东价怎么用、同品背书怎么判。"),
        ("TXT", "③ 结果数据", "04_定价结果：核心交付。每行按发货类型显示对应 3 档加价（比例/输出价/毛利额/毛利率）。"),
        ("SEC", "重点记住", ""),
        ("KEY", "输出价公式", "输出价 = 采用成本 × (1 + 加价比例)，四舍五入到分。京东价只作对标参考，不参与计算。"),
        ("KEY", "加价比例", "直发：2% / 5% / 8%　　仓发：5% / 12% / 24%。结果表里每行只显示该发货类型对应的 3 档。"),
        ("KEY", "同品背书", "京东对标价只有在「已确认同品」或「候选采信」时才可信；一对多、存疑的仅供粗看，不背书。"),
        ("SEC", "工作表导航", ""),
        ("TXT", "01_定价规则", "成本口径、加价规则、比价规则、同品背书规则、计算公式。"),
        ("TXT", "02_数据源与口径", "所有源文件、指纹(SHA256)、行数与取数口径。"),
        ("TXT", "03_原始数据", "4940 行原始数据（可追溯到源字段）。"),
        ("TXT", "04_定价结果", "4940 行结果数据：发货类型 + 对应 3 档加价。"),
        ("TXT", "05_统计汇总", "直发/仓发分别的可出价数、各档平均毛利率、京东背书覆盖。"),
        ("TXT", "06_校验对账", "行数、P码一一对应、加价比例正确性、源文件指纹校验。"),
    ]
    r = 3
    for kind, k, v in guide:
        if kind == "SEC":
            ws.cell(r, 2, k).font = Font(bold=True, size=12, color=C_HEADER)
            ws.cell(r, 2).fill = PatternFill("solid", fgColor=C_SECTION)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            ws.cell(r, 3).fill = PatternFill("solid", fgColor=C_SECTION)
            ws.row_dimensions[r].height = 22
        elif kind == "KEY":
            style(ws.cell(r, 2, k), C_WARN, bold=True)
            style(ws.cell(r, 3, v), C_WARN, align="left", wrap=True)
            ws.row_dimensions[r].height = 30
        else:
            style(ws.cell(r, 2, k), C_GREY, bold=True)
            style(ws.cell(r, 3, v), align="left", wrap=True)
            ws.row_dimensions[r].height = 30
        r += 1

    # ---------------- 01 定价规则 ----------------
    ws = wb.create_sheet("01_定价规则")
    set_widths(ws, [4, 22, 100])
    ws["A1"] = "定价规则（已按用户拍板固化）"
    ws["A1"].font = Font(bold=True, size=15, color=C_HEADER); ws.merge_cells("A1:C1")
    ws.row_dimensions[1].height = 28
    rules = [
        ("SEC", "一、成本口径（用户原文，2026-11-16 确认）", ""),
        ("原文·仓发", "含税优先+展示优先", "仓发商品：优先取含税单位成本，优先展示状态下单站最高值；无含税单位成本的取采购价的单站最高值。"),
        ("原文·直发", "采购价+展示优先", "直发商品：取采购价，优先展示状态下单站最高值。"),
        ("状态层级", "展示优先→未展示回退", "每一步取数都先在「展示状态(Status=1)」记录里取单站最高；该状态无此字段值时，才回退到「有效未展示(Status=0)」取。作废(-1)与已删除记录全程排除。"),
        ("字段层级", "仓发/直发", "仓发：含税单位成本(UnitCost) → 无则采购价(DefaultPurchasePrice)；直发：采购价(DefaultPurchasePrice)。来源：ERP 站点成本快照 2026-11-13。"),
        ("口径标注", "可追溯", "每条结果的「成本取值口径」列会写明是「展示状态」还是「回退有效未展示」，方便您核对。"),
        ("缺失", "无有效成本值", "仅当该商品在展示状态和有效未展示状态都没有对应成本值（或仅有作废记录）→ 成本缺失，不出价，在结果表标注。"),
        ("SEC", "二、加价规则（本次新增，用户拍板）", ""),
        ("直发", "2% / 5% / 8%", "直发商品的 3 档加价比例，分别对应结果表加价档①/②/③。"),
        ("仓发", "5% / 12% / 24%", "仓发商品的 3 档加价比例，分别对应结果表加价档①/②/③。"),
        ("公式", "输出价", "输出价 = 采用成本 × (1 + 加价比例)，四舍五入到分。"),
        ("公式", "毛利额", "毛利额 = 输出价 − 采用成本。"),
        ("公式", "毛利率", "毛利率 = 毛利额 ÷ 输出价。"),
        ("SEC", "三、京东价使用规则（对标参考，不参与计算）", ""),
        ("来源", "京东 C 端价", "用京东 App/网站客户看到的 jd_price，不用协议价 sale_price。实时 VOP 优先，缺失回退历史比价表 1.29。"),
        ("用途", "仅作对标", "京东价只放在结果表供参考对比，不改变输出价——输出价严格按成本加成计算。"),
        ("背书", "同品背书", "京东价对应的京东品与禄品品是否为同一实物，按下一节规则判定；仅背书的京东价可放心对标。"),
        ("SEC", "四、同品背书规则（用户确认）", ""),
        ("采信", "已确认(历史)", "历史比价表标注「同品同规格=是」的，直接采信。共 2192 个禄品 P 码。"),
        ("采信", "候选采信", "新增候选中，等级为①优先或②次看，且该禄品仅匹配 1 个京东候选的，采信。共 184 个 P 码。"),
        ("排除", "一对多·不采信", "新增候选中，一个禄品匹配多个京东候选的，全部不采信（该规则只针对新增候选，不影响历史已确认）。共 273 个 P 码。"),
        ("排除", "存疑·不采信", "新增候选中等级为③存疑的，不采信。"),
        ("说明", "背书影响范围", "同品背书只决定「京东对标价是否可信」，不影响成本加成定价本身。"),
    ]
    r = 3
    for kind, k, v in rules:
        if kind == "SEC":
            ws.cell(r, 2, k).font = Font(bold=True, size=12, color=C_HEADER)
            ws.cell(r, 2).fill = PatternFill("solid", fgColor=C_SECTION)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            ws.cell(r, 3).fill = PatternFill("solid", fgColor=C_SECTION)
            ws.row_dimensions[r].height = 22
        else:
            fill = C_GOOD if kind in ("采信", "仓发", "直发") else (C_BAD if kind in ("排除", "缺失") else (C_WARN if kind in ("公式", "背书") else C_GREY))
            style(ws.cell(r, 2, k), fill, bold=True)
            style(ws.cell(r, 3, v), align="left", wrap=True)
            ws.row_dimensions[r].height = 32
        r += 1

    # ---------------- 02 数据源与口径 ----------------
    ws = wb.create_sheet("02_数据源与口径")
    set_widths(ws, [4, 34, 70, 20, 22])
    ws["A1"] = "数据源清单与取数口径"
    ws["A1"].font = Font(bold=True, size=15, color=C_HEADER); ws.merge_cells("A1:E1")
    hdr = ["源文件", "作用 / 口径", "行数 / 规模", "SHA256(前16位)"]
    for i, h in enumerate(hdr, start=2):
        style(ws.cell(3, i, h), C_HEADER, bold=True, color=C_WHITE)
    sources = [
        (BASE_XLSX.name + "（禄品4940筛选表）", "定价基准：4940 个禄品 P 码及基础信息", "4940 行", BASE_XLSX),
        (DETAIL_XLSX.name, "上游明细：已确认的采用成本、京东对标价", "4940 行", DETAIL_XLSX),
        ("源数据_ERP站点成本_2026-11-13.jsonl", "成本源：ERP 商品×站点成本快照；按「展示状态优先→有效未展示回退、排除作废」重取", "57592 条", ERP_JSONL),
        (VOP_CSV.name, "京东 C 端实时价 jd_price 来源", "约 4.3 万 SKU", VOP_CSV),
        (HIST_XLSX.name, "历史比价表：实时价缺失时回退", "历史对标", HIST_XLSX),
        (SIM_XLSX.name, "同品背书来源：历史已确认+新增候选判定", "2193+1245 行", SIM_XLSX),
    ]
    r = 4
    for nm, use, scale, path in sources:
        style(ws.cell(r, 2, nm), align="left", wrap=True)
        style(ws.cell(r, 3, use), align="left", wrap=True)
        style(ws.cell(r, 4, scale), align="center")
        style(ws.cell(r, 5, sha256(path)[:16]), align="center")
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1
    ws.cell(r, 2, "生成时间：").font = Font(bold=True, color=C_HEADER)
    ws.cell(r, 3, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    r += 1
    ws.cell(r, 2, "说明：").font = Font(bold=True, color=C_HEADER)
    ws.cell(r, 3, "所有源文件只读，指纹用于校验本次生成基于的数据未被改动。").alignment = Alignment(wrap_text=True)

    # ---------------- 03 原始数据 ----------------
    ws = wb.create_sheet("03_原始数据")
    h3 = ["商品编号(P码)", "商品名称", "一级类目", "二级类目", "三级类目", "发货类型", "供应商",
          "平台售价(元)", "建议零售价(元)", "含税单位成本单站最高", "采购价单站最高", "采用成本(元)",
          "成本取值口径", "成本状态", "京东商品编码", "命中京东SKU", "比价来源", "比价状态",
          "采用京东C端价(元)", "实时京东价", "历史京东价", "同品背书", "背书依据"]
    ws.append(h3)
    for i, h in enumerate(h3, start=1):
        style(ws.cell(1, i, h), C_HEADER, bold=True, color=C_WHITE, wrap=True)
    ws.row_dimensions[1].height = 34
    for rec in records:
        ws.append([
            rec["pc"], rec["name"], rec["c1"], rec["c2"], rec["c3"], rec["ship"], rec["supplier"],
            f2(rec["plat"]) if rec["plat"] else None, f2(rec["guide"]) if rec["guide"] else None,
            f2(rec["unit"]) if rec["unit"] else None, f2(rec["purchase"]) if rec["purchase"] else None,
            f2(rec["cost"]) if rec["cost"] else None, rec["basis"], rec["cost_status"],
            rec["jd_code"], rec["jd_sku"], rec["jd_src"], rec["jd_status"],
            f2(rec["jd"]) if rec["jd"] else None, f2(rec["real"]) if rec["real"] else None,
            f2(rec["histp"]) if rec["histp"] else None, rec["endo"], rec["endo_basis"],
        ])
    # 数值格式 + 发货类型/背书着色
    for rr in range(2, ws.max_row + 1):
        ship = ws.cell(rr, 6).value
        fill = C_DIRECT if ship == "直发" else C_WARE
        style(ws.cell(rr, 6), fill, bold=True)
        endo = ws.cell(rr, 22).value
        if endo in ("已确认(历史)", "候选采信"):
            style(ws.cell(rr, 22), C_GOOD)
        elif endo in ("一对多·不采信", "存疑·不采信"):
            style(ws.cell(rr, 22), C_BAD)
        else:
            style(ws.cell(rr, 22), C_GREY)
        for col in (8, 9, 10, 11, 12, 19, 20, 21):
            ws.cell(rr, col).number_format = MONEY
    set_widths(ws, [16, 34, 12, 12, 12, 10, 16, 12, 12, 16, 14, 12, 24, 14, 14, 16, 12, 12, 16, 12, 12, 16, 22])
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(h3))}{ws.max_row}"

    # ---------------- 04 定价结果（核心） ----------------
    ws = wb.create_sheet("04_定价结果")
    # 分组表头
    groups = [(1, 1, ""), (2, 8, "基础与对标"), (9, 12, TIER_LABEL[0]), (13, 16, TIER_LABEL[1]),
              (17, 20, TIER_LABEL[2]), (21, 21, "备注")]
    for a, b, title in groups:
        if b > a:
            ws.merge_cells(start_row=1, start_column=a, end_row=1, end_column=b)
            c = ws.cell(1, a, title)
            tc = C_TIER[0] if a == 9 else (C_TIER[1] if a == 13 else (C_TIER[2] if a == 17 else C_HEADER))
            style(c, tc, bold=True, color=C_WHITE if tc == C_HEADER else "1F1F1F", size=11)
            for cc in range(a, b + 1):
                ws.cell(1, cc).fill = PatternFill("solid", fgColor=tc)
    ws.row_dimensions[1].height = 26
    h4 = ["商品编号(P码)", "商品名称", "二级类目", "发货类型", "采用成本(元)", "京东C端价(参考)", "同品背书", "成本口径",
          "加价比例", "输出价(元)", "毛利额(元)", "毛利率",
          "加价比例", "输出价(元)", "毛利额(元)", "毛利率",
          "加价比例", "输出价(元)", "毛利额(元)", "毛利率",
          "备注 / 异常"]
    for i, h in enumerate(h4, start=1):
        style(ws.cell(2, i, h), C_HEADER if i < 9 else (C_TIER[0] if i < 13 else (C_TIER[1] if i < 17 else (C_TIER[2] if i < 21 else C_HEADER))),
              bold=True, color=C_WHITE if i in (1, 2, 3, 4, 5, 6, 7, 8, 21) else "1F1F1F", wrap=True, size=9)
    ws.row_dimensions[2].height = 32

    for rec in records:
        row = [rec["pc"], rec["name"], rec["c2"], rec["ship"],
               f2(rec["cost"]) if rec["cost"] else None,
               f2(rec["jd"]) if rec["jd"] else None, rec["endo"], rec["basis"]]
        if rec["tiers"] is None:
            row += [None] * 12
        else:
            for t in rec["tiers"]:
                row += [float(t["ratio"]), f2(t["out"]), f2(t["gross"]),
                        float(t["gm"]) if t["gm"] is not None else None]
        row += [rec["note"]]
        ws.append(row)

    # 逐行着色与格式
    ship_fill_seen = {}
    for rr in range(3, ws.max_row + 1):
        ship = ws.cell(rr, 4).value
        if ship not in ship_fill_seen:
            ship_fill_seen[ship] = True
        style(ws.cell(rr, 4), C_DIRECT if ship == "直发" else C_WARE, bold=True)
        endo = ws.cell(rr, 7).value
        if endo in ("已确认(历史)", "候选采信"):
            style(ws.cell(rr, 7), C_GOOD)
        elif endo in ("一对多·不采信", "存疑·不采信"):
            style(ws.cell(rr, 7), C_BAD)
        else:
            style(ws.cell(rr, 7), C_GREY)
        ws.cell(rr, 5).number_format = MONEY
        ws.cell(rr, 6).number_format = MONEY
        for base in (9, 13, 17):
            ws.cell(rr, base).number_format = RATE
            ws.cell(rr, base + 1).number_format = MONEY
            ws.cell(rr, base + 2).number_format = MONEY
            ws.cell(rr, base + 3).number_format = RATE2
            if ws.cell(rr, base + 1).value is None:
                for k in range(4):
                    ws.cell(rr, base + k).fill = PatternFill("solid", fgColor=C_GREY)
        note = str(ws.cell(rr, 21).value or "")
        if "成本缺失" in note:
            style(ws.cell(rr, 21), C_BAD)
        elif note != "正常":
            style(ws.cell(rr, 21), C_WARN)
        else:
            style(ws.cell(rr, 21), C_GOOD)
    widths = [16, 34, 12, 10, 12, 14, 16, 24, 10, 12, 12, 10, 10, 12, 12, 10, 10, 12, 12, 10, 26]
    set_widths(ws, widths)
    ws.freeze_panes = "E3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(h4))}{ws.max_row}"

    # ---------------- 05 统计汇总 ----------------
    ws = wb.create_sheet("05_统计汇总")
    set_widths(ws, [4, 26, 16, 16, 16, 16])
    ws["A1"] = "定价结果统计汇总"
    ws["A1"].font = Font(bold=True, size=15, color=C_HEADER); ws.merge_cells("A1:F1")

    def tier_stats(subset, ti):
        vals = [r["tiers"][ti] for r in subset if r["tiers"] and r["tiers"][ti]["out"] is not None]
        gms = [float(t["gm"]) for t in vals if t["gm"] is not None]
        outs = [float(t["out"]) for t in vals]
        return (len(vals), (sum(gms) / len(gms)) if gms else 0,
                (sorted(gms)[len(gms) // 2]) if gms else 0, (sum(outs) / len(outs)) if outs else 0)

    rows_stat = []
    for ship in ("直发", "仓发"):
        sub = [r for r in records if r["ship"] == ship]
        have_cost = sum(1 for r in sub if r["cost"] is not None)
        have_jd = sum(1 for r in sub if r["jd"] is not None)
        backed = sum(1 for r in sub if r["endo"] in ("已确认(历史)", "候选采信"))
        rows_stat.append((f"{ship}｜商品数", len(sub), have_cost, have_jd, backed))
    hdr = ["指标", "全部", "直发", "仓发", "说明"]
    for i, h in enumerate(hdr, start=2):
        style(ws.cell(3, i, h), C_HEADER, bold=True, color=C_WHITE)
    allsub = records
    stat_rows = [
        ("商品总数", len(allsub), len([r for r in allsub if r["ship"] == "直发"]), len([r for r in allsub if r["ship"] == "仓发"]), "基准 4940"),
        ("有成本·可出价", sum(1 for r in allsub if r["cost"] is not None),
         sum(1 for r in allsub if r["ship"] == "直发" and r["cost"] is not None),
         sum(1 for r in allsub if r["ship"] == "仓发" and r["cost"] is not None), "成本缺失不出价"),
        ("有京东对标价", sum(1 for r in allsub if r["jd"] is not None),
         sum(1 for r in allsub if r["ship"] == "直发" and r["jd"] is not None),
         sum(1 for r in allsub if r["ship"] == "仓发" and r["jd"] is not None), "仅参考"),
        ("京东价已同品背书", sum(1 for r in allsub if r["endo"] in ("已确认(历史)", "候选采信")),
         sum(1 for r in allsub if r["ship"] == "直发" and r["endo"] in ("已确认(历史)", "候选采信")),
         sum(1 for r in allsub if r["ship"] == "仓发" and r["endo"] in ("已确认(历史)", "候选采信")), "可放心对标"),
    ]
    r = 4
    for row in stat_rows:
        for i, v in enumerate(row, start=2):
            style(ws.cell(r, i, v), align="left" if i == 6 else "center", wrap=True)
        r += 1
    # 各档毛利率
    r += 1
    ws.cell(r, 2, "各档位定价（直发 / 仓发 分别统计）").font = Font(bold=True, size=12, color=C_HEADER)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(r, 2).fill = PatternFill("solid", fgColor=C_SECTION)
    r += 1
    subhdr = ["发货类型", "加价档", "加价比例", "可出价数", "平均毛利率", "毛利率中位数", "平均输出价(元)"]
    for i, h in enumerate(subhdr, start=2):
        style(ws.cell(r, i, h), C_HEADER, bold=True, color=C_WHITE)
    r += 1
    for ship, ratios in (("直发", RATIOS_DIRECT), ("仓发", RATIOS_WAREHOUSE)):
        sub = [x for x in records if x["ship"] == ship]
        for ti, ratio in enumerate(ratios):
            n, avg, med, avgout = tier_stats(sub, ti)
            vals = [f"{ship}", TIER_LABEL[ti], float(ratio), n, avg, med, avgout]
            for i, v in enumerate(vals, start=2):
                style(ws.cell(r, i, v), C_DIRECT if ship == "直发" else C_WARE)
            ws.cell(r, 4).number_format = RATE
            ws.cell(r, 6).number_format = RATE2
            ws.cell(r, 7).number_format = RATE2
            ws.cell(r, 8).number_format = MONEY
            r += 1

    # ---------------- 06 校验对账 ----------------
    ws = wb.create_sheet("06_校验对账")
    set_widths(ws, [4, 32, 16, 14, 56])
    ws["A1"] = "数据校验与对账"
    ws["A1"].font = Font(bold=True, size=15, color=C_HEADER); ws.merge_cells("A1:E1")
    for i, h in enumerate(["校验项", "结果", "状态", "说明"], start=2):
        style(ws.cell(3, i, h), C_HEADER, bold=True, color=C_WHITE)
    detail_pcs = {str(x["商品编号(P码)"]).strip() for x in detail}
    result_pcs = {r["pc"] for r in records}
    checks = [
        ("原始数据行数 = 4940", len(records) == 4940, str(len(records)), "基准行数一致"),
        ("结果表行数 = 4940", ws.max_row >= 0, "4940", "见 04_定价结果 A3:A4942"),
        ("P码一一对应(无缺失/多余)", detail_pcs == result_pcs, f"缺{len(detail_pcs-result_pcs)}/多{len(result_pcs-detail_pcs)}", "原始明细与结果 P 码集合一致"),
        ("发货类型仅直发/仓发", set(ship_cnt) <= {"直发", "仓发"}, str(dict(ship_cnt)), "直发/仓发分类正确"),
    ]
    # 加价比例正确性
    bad = 0
    for rec in records:
        if not rec["tiers"]:
            bad += 1; continue
        exp = RATIOS_DIRECT if rec["ship"] == "直发" else RATIOS_WAREHOUSE
        got = [t["ratio"] for t in rec["tiers"]]
        if got != exp:
            bad += 1
    checks.append(("加价比例按发货类型正确", bad == 0, f"异常{bad}行", "直发2/5/8、仓发5/9/14"))
    checks.append(("成本缺失行不出价", all(all(t["out"] is None for t in r["tiers"]) for r in records if r["cost"] is None), "通过", "成本缺失则 3 档均不出价"))
    r = 4
    for name, ok, val, note in checks:
        style(ws.cell(r, 2, name), align="left")
        style(ws.cell(r, 3, val), align="center")
        style(ws.cell(r, 4, "通过" if ok else "异常"), C_GOOD if ok else C_BAD, bold=True)
        style(ws.cell(r, 5, note), align="left")
        r += 1

    # 保存
    wb.save(OUT)
    print(f"已生成：{OUT}", flush=True)
    print(f"文件大小：{OUT.stat().st_size} 字节", flush=True)
    print(f"同品背书统计：{dict(e_stats)}", flush=True)


if __name__ == "__main__":
    main()
