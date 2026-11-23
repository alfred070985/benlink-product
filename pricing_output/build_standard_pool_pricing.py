#!/usr/bin/env python3
"""生成分销标准池全量定价表——技术公司视角，供货给所有客户。

口径（用户 2026-11-19 拍板）：
  - 池子：分销池(ProductPoolType=1) 北京/广州/上海/武汉(1,2,3,7) 四站全上架且有库存
  - 成本：ERP 快照按「展示优先→未展示回退、排除作废」；直发采购价 / 仓发含税优先，单站最高
  - 发货类型：成本明细中 OriginDirectDelivery=True → 直发，否则仓发（禄品池标定 99%/100% 吻合）
  - 加价档：直发 2/5/8%，仓发 5/12/24%；输出价=成本×(1+比例)
  - 对标价：相似度≥0.75 自动采信的京东VOP价；未匹配用 ERP 基准价(basicPrice)兜底并标注
  - 建议档：默认中间档②；档②空间率<10%→降档①；档②空间率≥30% 且档③仍≥10%→升档③
  - 全量给价，不做"选/不选"筛选

只读源文件，输出全新工作簿。
"""
import json
import os
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import sys as _sys
_sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_similarity import parse_pack

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")

POOL = Path(os.environ.get("PRICING_POOL", PRICING_DIR / f"标准池_四站全上架有库存_{TODAY}.jsonl"))
INFO = Path(os.environ.get("PRICING_INFO", PRICING_DIR / f"源数据_标准池商品信息_{TODAY}.jsonl"))
COST = Path(os.environ.get("PRICING_COST", PRICING_DIR / f"源数据_标准池ERP站点成本_{TODAY}.jsonl"))
MATCH = Path(os.environ.get("PRICING_MATCH", PRICING_DIR / f"标准池京东相似匹配_{TODAY}.jsonl"))
OUT = PRICING_DIR / f"分销标准池定价_{TODAY}.xlsx"

# 天镜智价快照（取目录下最新一份；京东对标价统一走天镜，VOP CSV 仅作天镜无价时的备选）
import glob as _glob
def _latest(pat):
    fs = sorted(_glob.glob(str(PRICING_DIR / pat)))
    return Path(fs[-1]) if fs else None
HIST_JSONL = _latest("源数据_天镜价格历史全量_*.jsonl")
MAP_JSONL = _latest("源数据_天镜同品映射_*.jsonl")

RATIOS_DIRECT = [Decimal("0.02"), Decimal("0.05"), Decimal("0.08")]
RATIOS_WAREHOUSE = [Decimal("0.05"), Decimal("0.12"), Decimal("0.24")]
HI_THRESHOLD = Decimal("0.30")   # 档②空间率≥30% 视为毛利空间巨大(建议值,待用户确认)
SPACE_FLOOR = Decimal("0.10")

C_HEADER = "1F4E79"
C_WHITE = "FFFFFF"
C_DIRECT = "FCE4D6"
C_WARE = "DDEBF7"
C_NEW = "E2EFDA"
C_GREY = "F2F2F2"
C_SUGGEST = "FFF2CC"


def d(x):
    if x in (None, "", 0, "0"):
        return None
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    return v if v > 0 else None


def load_costs():
    recs = defaultdict(list)
    with COST.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                o = json.loads(line)
                recs[o["ProductBasicSysNo"]].append(o)
    return recs


def load_tianjing():
    """加载天镜历史价库(每SKU最新一条) + 天镜人工核实映射(按P码数字聚合)。

    返回 (hist: {sku: (price,time)}, same: {pbasic: [sku...]},
          diff: {pbasic: [(sku,ratio)...]}, sim: {pbasic: [sku...]})。
    """
    hist = {}
    if HIST_JSONL and HIST_JSONL.exists():
        with HIST_JSONL.open(encoding="utf-8") as f:
            for line in f:
                if not line.strip():
                    continue
                o = json.loads(line)
                sku = str(o.get("skuId", "")).strip()
                if not sku or sku in hist or o.get("status") != "success":
                    continue
                price = d(o.get("jdPrice")) or d(o.get("salePrice"))
                if price is None:
                    continue
                hist[sku] = price
    same, diff, sim = defaultdict(list), defaultdict(list), defaultdict(list)
    if MAP_JSONL and MAP_JSONL.exists():
        with MAP_JSONL.open(encoding="utf-8") as f:
            for line in f:
                if not line.strip():
                    continue
                o = json.loads(line)
                ic = str(o.get("internal_code") or "").strip()
                if not (len(ic) > 1 and ic[0].upper() == "P" and ic[1:].isdigit()):
                    continue
                pb = int(ic[1:])
                vs = str(o.get("vop_sku") or "").strip()
                if not vs:
                    continue
                rt = o.get("relation_type")
                if rt == "same_sku":
                    same[pb].append(vs)
                elif rt == "diff_spec":
                    diff[pb].append((vs, o.get("ratio") or ""))
                else:
                    sim[pb].append(vs)
    return hist, same, diff, sim


def pick(field, rows):
    vals = [r.get(field) for r in rows if isinstance(r.get(field), (int, float)) and r.get(field) > 0]
    return max(vals) if vals else None


def resolve_cost(rs):
    """返回 (成本, 口径, 是否直发)。沿用拍板成本口径 + 发货标定规则。"""
    show = [r for r in rs if r.get("Status") == 1 and not r.get("IsDelete")]
    unshow = [r for r in rs if r.get("Status") == 0 and not r.get("IsDelete")]
    direct = any(r.get("OriginDirectDelivery") for r in rs)

    def pick_layer(field, label):
        v = pick(field, show)
        if v:
            return v, f"{label}·展示状态单站最高"
        v = pick(field, unshow)
        if v:
            return v, f"{label}·回退有效未展示单站最高"
        return None, ""

    if direct:
        v, basis = pick_layer("DefaultPurchasePrice", "采购价")
        if v:
            return v, basis, True
        return None, "无有效采购价", True
    v, basis = pick_layer("UnitCost", "含税单位成本")
    if v:
        return v, basis, False
    v, basis = pick_layer("DefaultPurchasePrice", "采购价")
    if v:
        return v, basis + "(无含税单位成本)", False
    return None, "有效状态记录无成本", False


def price_tiers(cost, direct):
    ratios = RATIOS_DIRECT if direct else RATIOS_WAREHOUSE
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


def suggest(tiers, jd, is_fallback):
    """新选档规则：默认中档；档②空间<10%降档①；档②空间≥30%且档③≥10%升档③。
    兜底价非市场价、空间率失真：不升档，仅空间<10%时降档①。"""
    if jd is None:
        return 1, "无对标价·默认中档"
    jd = Decimal(str(jd))
    spaces = [(Decimal(str(t[1] and (jd - Decimal(str(t[1]))) / jd)) if t[1] else None) for t in tiers]
    s2 = spaces[1]
    if s2 is None:
        return 1, "无对标价·默认中档"
    if s2 < SPACE_FLOOR:
        return 0, f"档②空间{s2:.0%}<10%·降档①"
    if is_fallback:
        return 1, f"兜底价空间失真·保守中档(表面空间{s2:.0%})"
    if s2 >= HI_THRESHOLD and spaces[2] is not None and spaces[2] >= SPACE_FLOOR:
        return 2, f"档②空间{s2:.0%}≥30%·毛利空间大·升档③"
    return 1, f"档②空间{s2:.0%}·中档"


def main():
    pool = {json.loads(l)["productSysNo"]: json.loads(l) for l in POOL.open(encoding="utf-8") if l.strip()}
    info = {json.loads(l)["productSysNo"]: json.loads(l) for l in INFO.open(encoding="utf-8") if l.strip()}
    by_basic = defaultdict(list)
    for psys, p in pool.items():
        by_basic[p["productBasicSysNo"]].append(psys)
    costs = load_costs()
    match = {}
    for l in MATCH.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            if o.get("采纳"):
                match[o["productSysNo"]] = o

    print(f"标准池 {len(pool)} | 商品信息 {len(info)} | 采信对标 {len(match)}", flush=True)

    hist, same, diff, sim = load_tianjing()
    print(f"天镜价库SKU {len(hist)} | 同款映射P码 {len(same)} | 不同规 {len(diff)} | 相似 {len(sim)}", flush=True)

    records = []
    stat = Counter()
    for psys in sorted(pool):
        p = pool[psys]
        inf = info.get(psys, {})
        cost, basis, direct = resolve_cost(costs.get(p["productBasicSysNo"], []))
        pb = p["productBasicSysNo"]
        m = match.get(psys)
        jd = jd_src = jd_sku = None
        # 优先级1：天镜人工核实·同款(same_sku) → 天镜价，无价回退VOP
        skus = same.get(pb) or []
        tj = next((hist[s] for s in skus if s in hist), None)
        if tj is not None:
            jd, jd_src, jd_sku = tj, "天镜智价·人工核实同款", skus[0]
            stat["对标_天镜同款"] += 1
        elif skus and m:
            jd, jd_src, jd_sku = m["jd_price"], "京东VOP·同款映射备选价", m["jd_sku"]
            stat["对标_VOP备选"] += 1
        # 优先级2：天镜diff_spec → 折算（用户拍板：比例=京东侧件数:本Life件数；
        # 单件价=京东价÷分子；本Life件数优先取商品名解析，与禄品池同套规格解析逻辑）
        elif diff.get(pb):
            vs, ratio = diff[pb][0]
            tj = hist.get(vs)
            if tj is not None:
                try:
                    a, b = (int(x) for x in str(ratio).split(":"))
                    if a <= 0 or b <= 0:
                        raise ValueError
                except Exception:
                    a, b = 1, 1
                cnt = (parse_pack(inf.get("name") or "").get("count") or 1)
                jd = round(tj / a * cnt, 2)
                jd_src = f"天镜智价·不同规折算({ratio}·{cnt}件)"
                jd_sku = vs
                stat["对标_天镜不同规"] += 1
            else:
                stat["待确认_不同规无价"] += 1
        elif sim.get(pb):
            stat["仅参考_相似关系"] += 1
        # 优先级3：算法相似采信(≥0.75) → 天镜价优先，VOP备选
        elif m:
            tj = hist.get(m["jd_sku"])
            if tj is not None:
                jd, jd_src, jd_sku = tj, f"天镜智价·相似采信({m['score']})", m["jd_sku"]
                stat["对标_相似采信天镜"] += 1
            else:
                jd, jd_src, jd_sku = m["jd_price"], f"京东VOP·相似采信({m['score']})", m["jd_sku"]
                stat["对标_相似采信VOP"] += 1
        # 兜底：ERP基准价
        if jd is None:
            bp = p.get("basicPrice") or {}
            vals = [v for v in bp.values() if isinstance(v, (int, float)) and v > 0]
            if vals:
                jd, jd_src = sum(vals) / len(vals), "ERP基准价兜底·非京东价"
                jd_sku = ""
                tag = "天镜diff_spec待确认" if diff.get(pb) else ("天镜similar仅参考" if sim.get(pb) else "")
                stat["对标_ERP兜底"] += 1
            else:
                jd, jd_src, jd_sku = None, "无对标价", ""
                stat["对标_无"] += 1
        if cost is None:
            stat["成本缺失"] += 1
        tiers = price_tiers(cost, direct)
        idx, reason = suggest(tiers, jd, jd_src.startswith("ERP基准价"))
        records.append({
            "psys": psys, "pbasic": pb,
            "name": inf.get("name", ""), "c1": inf.get("c1"), "c2": inf.get("c2"), "c3": inf.get("c3"),
            "ship": "直发" if direct else "仓发", "cost": cost, "basis": basis,
            "jd": jd, "jd_src": jd_src, "jd_sku": jd_sku,
            "tiers": tiers, "suggest_idx": idx, "suggest_reason": reason,
        })
        stat["直发" if direct else "仓发"] += 1

    build_workbook(records, stat)

    # 统计落盘：供 build_chain_report.py 生成链路报告用（周频复跑，数字每周滚动）
    import statistics
    ship_stat = {"直发": stat.get("直发", 0), "仓发": stat.get("仓发", 0)}
    suggest_stat = defaultdict(Counter)
    spaces = {"真实对标": [], "ERP兜底": []}
    chk_cells = chk_err = 0
    ratios_of = {"直发": [0.02, 0.05, 0.08], "仓发": [0.05, 0.12, 0.24]}
    for r in records:
        is_real = "兜底" not in r["jd_src"] and r["jd_src"] != "无对标价"
        key = "真实对标" if is_real else "ERP兜底"
        suggest_stat[key][f"档{r['suggest_idx']+1}"] += 1
        if r["jd"] and r["tiers"][r["suggest_idx"]][1]:
            sp = (r["jd"] - r["tiers"][r["suggest_idx"]][1]) / r["jd"]
            spaces[key].append(sp)
        if r["cost"] is not None:
            ratios = ratios_of[r["ship"]]
            for t in range(3):
                chk_cells += 1
                if abs(r["tiers"][t][1] - round(r["cost"] * (1 + ratios[t]), 2)) > 0.011:
                    chk_err += 1
    stats = {
        "日期": TODAY,
        "总行数": len(records),
        "发货类型": dict(ship_stat),
        "对标来源": {k.replace("对标_", ""): v for k, v in stat.items() if k.startswith("对标_")},
        "天镜特殊": {"diff_spec无价未折算": stat.get("待确认_不同规无价", 0),
                   "similar仅参考": stat.get("仅参考_相似关系", 0)},
        "成本缺失": stat.get("成本缺失", 0),
        "建议档分布": {k: dict(v) for k, v in suggest_stat.items()},
        "建议档空间率中位数": {
            "真实对标": round(statistics.median(spaces["真实对标"]), 4) if spaces["真实对标"] else None,
            "ERP兜底": round(statistics.median(spaces["ERP兜底"]), 4) if spaces["ERP兜底"] else None,
        },
        "公式校验": {"格数": chk_cells, "错误数": chk_err},
    }
    STATS_OUT = PRICING_DIR / f"定价统计_{TODAY}.json"
    STATS_OUT.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False, indent=2), flush=True)


def build_workbook(records, stat):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "01_定价结果"

    hdr = ["商品编号", "商品名称", "一级类目", "二级类目", "三级类目", "发货类型",
           "采用成本(元)", "成本口径", "对标价(元)", "对标来源", "对标SKU",
           "档①比例", "档①输出价", "档①毛利额", "档①毛利率",
           "档②比例", "档②输出价", "档②毛利额", "档②毛利率",
           "档③比例", "档③输出价", "档③毛利额", "档③毛利率",
           "建议档", "建议输出价(元)", "建议档空间率", "建议说明"]
    ws.append(hdr)
    for c in range(1, len(hdr) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = PatternFill("solid", fgColor=C_HEADER)
        cell.font = Font(color=C_WHITE, bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in records:
        t = r["tiers"]
        jd = r["jd"]
        s_price = t[r["suggest_idx"]][1]
        s_space = (jd - s_price) / jd if (jd and s_price) else None
        row = [r["psys"], r["name"], r["c1"], r["c2"], r["c3"], r["ship"],
               r["cost"], r["basis"], r["jd"], r["jd_src"], r["jd_sku"]]
        for i in range(3):
            row += [t[i][0], t[i][1], t[i][2], t[i][3]]
        row += [f"档{r['suggest_idx']+1}", s_price, s_space, r["suggest_reason"]]
        ws.append(row)

    widths = [10, 36, 8, 8, 8, 8, 10, 24, 10, 26, 10] + [8, 10, 10, 8] * 3 + [7, 12, 11, 26]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(hdr))}1"

    # 数字格式
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx in (7, 9, 13, 17, 21, 26):
            row[idx - 1].number_format = "0.00"
        for idx in (12, 14, 16, 18, 20, 22, 24, 27):
            row[idx - 1].number_format = "0.0%"
        for idx in (12, 16, 20):
            row[idx - 1].number_format = "0%"

    # 统计页
    ws2 = wb.create_sheet("02_统计汇总")
    ws2.append(["指标", "数值", "说明"])
    n = len(records)
    ws2.append(["总商品数", n, "四站全上架+四站有库存"])
    for k, v in stat.items():
        ws2.append([k, v, f"{v/n:.1%}"])
    ws2.append(["建议档分布", "", ""])
    sc = Counter(f"档{r['suggest_idx']+1}" for r in records)
    for k in ("档1", "档2", "档3"):
        ws2.append([k, sc.get(k, 0), f"{sc.get(k,0)/n:.1%}"])
    ws2.append([])
    ws2.append(["建议档规则", "默认中间档②；档②空间<10%降档①；档②空间≥30%且档③仍≥10%升档③", ""])
    ws2.append(["升档阈值30%", "建议值，待用户确认", ""])
    for c in range(1, 4):
        ws2.cell(row=1, column=c).fill = PatternFill("solid", fgColor=C_HEADER)
        ws2.cell(row=1, column=c).font = Font(color=C_WHITE, bold=True)
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 60
    ws2.column_dimensions["C"].width = 30

    wb.save(OUT)
    print(f"产出: {OUT}", flush=True)


if __name__ == "__main__":
    main()
