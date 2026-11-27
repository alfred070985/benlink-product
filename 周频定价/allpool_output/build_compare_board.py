#!/usr/bin/env python3
"""本来品 ↔ 京东VOP / 猫超 相似度比对盘（全量、每品top3候选）。

用户 2026-11-25 拍板：重做一份完整比对盘，主体只有本来品（纯池58,423，
已按 ERP t_goods_map 剔 16,644 京东品），只对照 VOP 和猫超，全量覆盖，
每品展示相似度最高的 3 个候选，带双方价格/规格折算/当前定价状态，
供人工审阅哪些边界候选可采信。

输出：交付物/2026-W48/03_货盘底表/5_跨源对齐/本来品_京东猫超相似度比对盘_2026-11-25.xlsx
  00_说明
  01_本来品x京东VOP   （每品 top3 候选一行）
  02_本来品x猫超      （每品 top3 候选一行）
  03_品级汇总          （每品一行：两源最佳候选 + 当前对标价状态）

路径全部走 latest()，遵守 R1 自包含。
"""
import csv
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "pricing_output"))

from latest import latest  # noqa: E402
from build_similarity import parse_pack  # noqa: E402

TODAY = datetime.now().strftime("%Y-%m-%d")
# 用户 2026-11-26 拍板：改用 v2 精排结果，采信阈值 0.8（原文件按 0.7 打的标不沿用，本脚本按 0.8 重判）
ACCEPT = 0.8
RESCUE_FLOOR = 0.75  # 边界可救下限；同时是高边界折算区间(0.75~0.80)下沿
OUT = (ROOT / "交付物/2026-W48/03_货盘底表/5_跨源对齐"
       / f"本来品_京东猫超相似度比对盘_v2精排_{TODAY}.xlsx")


def cnt_of(name):
    try:
        return (parse_pack(name or "").get("count") or 1) or 1
    except Exception:
        return 1


def cnt_from_attr(attr, name):
    """优先用 v2 精排链路已抽取的件数（修了佳沛“单果克重误当件数”的 bug），
    没有则回退到 parse_pack。"""
    try:
        c = (attr or {}).get("count")
        if c and int(c) > 0:
            return int(c)
    except (ValueError, TypeError):
        pass
    return cnt_of(name)


def load_names():
    """pbasic -> 本网品名"""
    info = latest("pricing_output/源数据_标准池商品信息_*.jsonl")
    pool = latest("pricing_output/标准池_四站全上架有库存_*.jsonl")
    p2b = {str(json.loads(l)["productSysNo"]): str(json.loads(l)["productBasicSysNo"])
           for l in pool.open(encoding="utf-8") if l.strip()}
    name = {}
    for l in info.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            pb = p2b.get(str(o["productSysNo"]))
            if pb:
                name[pb] = o.get("name") or ""
    return name


def load_bl_cost():
    import build_standard_pool_pricing as base
    base.POOL = latest("pricing_output/标准池_四站全上架有库存_*.jsonl", "PRICING_POOL")
    base.INFO = latest("pricing_output/源数据_标准池商品信息_*.jsonl", "PRICING_INFO")
    base.COST = latest("pricing_output/源数据_标准池ERP站点成本_*.jsonl", "PRICING_COST")
    costs = base.load_costs()
    out = {}
    for pb, rows in costs.items():
        c, _, _ = base.resolve_cost(rows)
        if c and c > 0:
            out[str(pb)] = c
    return out


def load_status():
    """pbasic -> 当前货盘一对标价状态（来自成品对标来源列）"""
    xl = latest("交付物/2026-W48/01_货盘一_本来定制/分销标准池定价_*.xlsx")
    pool = latest("pricing_output/标准池_四站全上架有库存_*.jsonl")
    p2b = {str(json.loads(l)["productSysNo"]): str(json.loads(l)["productBasicSysNo"])
           for l in pool.open(encoding="utf-8") if l.strip()}
    wb = openpyxl.load_workbook(xl, read_only=True, data_only=True)
    ws = wb["01_定价结果"]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    pi, si, hi = hdr.index("商品编号"), hdr.index("对标来源"), hdr.index("对标价(元)")
    out = {}
    for r in it:
        if not r or not r[pi]:
            continue
        pb = p2b.get(str(r[pi]).strip(), str(r[pi]).strip())
        src = str(r[si] or "")
        if "高边界" in src:
            st = "高边界折算"
        elif "向量" in src:
            st = "向量采信"
        elif "天镜" in src:
            st = "天镜核实"
        elif "ERP" in src:
            st = "ERP兜底(无真实价)"
        elif "京东VOP" in src:
            st = "旧映射价(其他)"
        else:
            st = src or "未知"
        out[pb] = (st, r[hi])
    wb.close()
    return out


def load_prices(which):
    """对方价格：京东 jd_price/sale_price/market_price；猫超 channel_price/coupon_price"""
    out = {}
    if which == "jd":
        v = latest("supply_pull/output/本来生活VOP_goods_本网ID.csv")
        for r in csv.DictReader(v.open(encoding="utf-8-sig")):
            def f(k):
                try:
                    x = float(r.get(k)) if r.get(k) else None
                    return x if x and x > 0 else None
                except (ValueError, TypeError):
                    return None
            out[str(r["sku_id"])] = {"channel": f("jd_price"), "hand": f("sale_price"),
                                     "market": f("market_price"), "name": r.get("sku_name", "")}
    else:
        m = latest("supply_pull/output/猫超价格_全量_*.jsonl")
        for l in m.open(encoding="utf-8"):
            if not l.strip():
                continue
            o = json.loads(l)
            def f(k):
                x = o.get(k)
                return x if isinstance(x, (int, float)) and x > 0 else None
            out[str(o["product_id"])] = {"channel": f("channel_price"), "hand": f("coupon_price"),
                                         "market": None, "name": ""}
    return out


def build(which, align_file):
    """返回 {pbasic: [候选dict,...]}，按相似度降序取 top3"""
    excl = {str(json.loads(l)["pbasic"]) for l in
            (ROOT / "交付物/2026-W48/03_货盘底表/5_跨源对齐/分销池_ERP京东品.jsonl").open(encoding="utf-8")
            if l.strip()}
    prices = load_prices(which)
    cands = defaultdict(list)
    for l in (HERE / align_file).open(encoding="utf-8"):
        if not l.strip():
            continue
        o = json.loads(l)
        a, b = str(o["a"]), str(o["b"])
        if a in excl or b not in prices:
            continue
        p = prices[b]
        ch = p["channel"]
        if not ch:
            # 到手价兜底：渠道价缺失时用券后/到手价
            ch = p["hand"]
        if not ch:
            continue
        cnt_bl = cnt_from_attr(o.get("a_attr"), o.get("a_name"))
        cnt_opp = cnt_from_attr(o.get("b_attr"), o.get("b_name"))
        if cnt_opp <= 0 or cnt_bl <= 0:
            continue
        adj = round(ch / cnt_opp * cnt_bl, 2)
        # v2 文件里的 tier 是按 0.7 打的标；本次用户改 0.8，统一按精排分重判，不沿用旧标
        tier = "采信" if o["sim"] >= ACCEPT else "边界"
        cands[a].append({
            "pbasic": a, "opp": b, "sim": o["sim"], "tier": tier,
            "rank": o["rank"], "conflict": o.get("conflict") or "",
            "opp_name": o.get("b_name") or p["name"],
            "channel": ch, "hand": p["hand"], "market": p["market"],
            "cnt_bl": cnt_bl, "cnt_opp": cnt_opp, "adj": adj,
        })
    out = {}
    for a, cs in cands.items():
        cs.sort(key=lambda x: -x["sim"])
        out[a] = cs[:3]
    return out


HDR_DETAIL = ["本来编号(P码)", "本来品名", "本来成本(元)", "当前定价状态", "当前对标价(元)",
              "对照货源", "对照编号", "对照品名", "精排分(v2)", "档位", "差异/冲突", "候选序",
              "对照渠道价(元)", "对照到手价(元)", "对照吊牌价(元)",
              "本来件数", "对照件数", "规格折算对标价(元)"]
HDR_SUM = ["本来编号(P码)", "本来品名", "本来成本(元)", "当前定价状态", "当前对标价(元)",
           "京东最佳精排分", "京东最佳品名", "京东折算对标价(元)", "京东档位",
           "猫超最佳精排分", "猫超最佳品名", "猫超折算对标价(元)", "猫超档位", "可救标记"]


def main():
    names = load_names()
    costs = load_bl_cost()
    status = load_status()
    pool = latest("pricing_output/标准池_四站全上架有库存_*.jsonl")
    pure = [str(json.loads(l)["productBasicSysNo"]) for l in pool.open(encoding="utf-8") if l.strip()]
    excl = {str(json.loads(l)["pbasic"]) for l in
            (ROOT / "交付物/2026-W48/03_货盘底表/5_跨源对齐/分销池_ERP京东品.jsonl").open(encoding="utf-8")
            if l.strip()}
    pure = [p for p in pure if p not in excl]
    print(f"纯本来池 {len(pure)} 品")

    jd = build("jd", "对齐_v2_bl_jd.jsonl")
    mc = build("mc", "对齐_v2_bl_mc.jsonl")
    print(f"有候选: 京东 {len(jd)} 品, 猫超 {len(mc)} 品")

    wb = openpyxl.Workbook(write_only=True)
    # ---- 00 说明 ----
    ws = wb.create_sheet("00_说明")
    intro = [
        ["本来品 ↔ 京东VOP / 猫超 相似度比对盘", ""],
        ["生成日期", TODAY],
        ["主体", f"纯本来池 {len(pure)} 品（75,067 − ERP t_goods_map 剔 16,644 京东品）"],
        ["对照货源", "京东VOP + 猫超（只有这两源，不含京东×猫超、东方甄选）"],
        ["模型(v2精排)", "bge-small 向量召回 top5 → 属性硬约束(品牌/无规格/超18倍) → BAAI/bge-reranker-v2-m3 精排分；每品展示最像的 3 个候选"],
        ["18倍规则(11-27)", "用户拍板：规格/件数/长度同维度差异>18倍→完全不当同品、直接排除(网线0.5米vs305米=610倍)；≤18倍算同款不同包装放行(水1瓶vs12瓶=12倍可比、vs24瓶=24倍排除)。旧版『件数/规格不等一律硬拦』已废除(误杀5L vs 2.5L等)。清洗营养含量/型号/单份克重后，69码真同款误伤0%"],
        ["规格矛盾减分", "本网明确单件 vs 对方明确多件(如珍珠米5KG vs 丝苗米5kg×4袋)判规格矛盾，精排分扣0.20后降档；对方不写规格(猫超普遍)视为信息缺失、不减分"],
        ["档位(阈值0.8)", "采信=减分后精排分≥0.80 且无硬约束冲突；边界=<0.80（原0.7打的标不沿用，本盘按0.8重判）"],
        ["阈值校准", "猫超 2,539 对真条码验证：0.8档真同款通过率83.9%、全量保留率57.7%；京东无真条码集，0.8借猫超校准间接推断"],
        ["规格折算价", "对照渠道价 ÷ 对照件数 × 本来件数；件数优先用 v2 抽取属性(已修佳沛单果克重误判)，回退 parse_pack"],
        ["高边界折算区间", "随采信线上移至 0.75~0.80（紧邻新采信线下方）；本盘只展示不实际折算，货盘重算时才应用"],
        ["当前定价状态", "取自货盘一成品对标来源列：天镜核实/向量采信/高边界折算/ERP兜底(无真实价)（货盘一未重算，仍为旧口径）"],
        ["可救标记", "当前 ERP兜底 且有边界候选（精排分≥0.75、非品牌冲突）→ 值得人工审阅是否采信"],
        ["用法", "按精排分排序看，勾中可采信的候选；属性冲突行在精排前已被拦截，故本盘冲突列多为空"],
    ]
    for row in intro:
        ws.append(row)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 90

    # ---- 明细 ----
    def detail_sheet(title, data, label):
        ws = wb.create_sheet(title)
        ws.append(HDR_DETAIL)
        n, covered = 0, 0
        for pb in pure:
            st, cur = status.get(pb, ("未知", None))
            for c in data.get(pb, []):
                ws.append([pb, names.get(pb, ""), costs.get(pb), st, cur, label,
                           c["opp"], c["opp_name"], c["sim"], c["tier"], c["conflict"],
                           c["rank"], c["channel"], c["hand"], c["market"],
                           c["cnt_bl"], c["cnt_opp"], c["adj"]])
                n += 1
            if pb in data:
                covered += 1
        return n, covered

    n_jd, cov_jd = detail_sheet("01_本来品x京东VOP", jd, "京东VOP")
    n_mc, cov_mc = detail_sheet("02_本来品x猫超", mc, "猫超")
    print(f"明细: 京东 {n_jd} 行({cov_jd}品), 猫超 {n_mc} 行({cov_mc}品)")

    # ---- 品级汇总 ----
    ws = wb.create_sheet("03_品级汇总")
    ws.append(HDR_SUM)
    for pb in pure:
        st, cur = status.get(pb, ("未知", None))
        bjd = jd.get(pb, [None])[0]
        bmc = mc.get(pb, [None])[0]
        def usable(cs):
            # 可救：边界、精排分≥0.75、且非品牌冲突（品牌冲突基本是误配）；看全部候选
            return any(c["tier"] == "边界" and c["sim"] >= RESCUE_FLOOR and c["conflict"] != "品牌冲突"
                       for c in cs)
        rescuable = "可救-待审" if (st.startswith("ERP兜底") and
                                  (usable(jd.get(pb, [])) or usable(mc.get(pb, [])))) else ""
        ws.append([
            pb, names.get(pb, ""), costs.get(pb), st, cur,
            bjd["sim"] if bjd else None, bjd["opp_name"] if bjd else "",
            bjd["adj"] if bjd else None, bjd["tier"] if bjd else "无候选",
            bmc["sim"] if bmc else None, bmc["opp_name"] if bmc else "",
            bmc["adj"] if bmc else None, bmc["tier"] if bmc else "无候选",
            rescuable,
        ])
    print(f"品级汇总 {len(pure)} 行")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"已生成: {OUT} ({OUT.stat().st_size / 1024 / 1024:.1f} MB)")


if __name__ == "__main__":
    main()
