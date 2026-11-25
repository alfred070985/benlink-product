#!/usr/bin/env python3
"""高边界候选（0.80-0.85 无硬约束冲突）规格折算后收敛为 1对1。

用户 2026-11-25 拍板：把 0.80-0.85 的高边界候选降档采信，按规格折算作对标价。
只处理本网侧候选（bl_jd / bl_mc），且本网品必须在纯池内（非被剔京东品）。

规格折算：对标品名件数 cnt_opp，本网品名件数 cnt_bl；
  折算对标价 = 对方渠道价 / cnt_opp * cnt_bl
件数解析复用 build_similarity.parse_pack（get count）。

输出：
  收敛_高边界_京东.jsonl  {pbasic, target, sim, price, src}
  收敛_高边界_猫超.jsonl  同结构
"""
import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "pricing_output"))

from latest import latest  # noqa: E402
from build_similarity import parse_pack  # noqa: E402

LO, HI = 0.80, 0.85
EXCLUDE = latest("allpool_output/分销池_ERP京东品.jsonl", "ERP_JD_EXCLUDE")


def load_names():
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
    """本网采用成本：用于折算合理性护栏（折算价/成本须落在合理区间）。"""
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


# 折算合理性护栏：对标价通常为成本的 0.5~8 倍；超出视为规格解析错误，丢弃
COST_RATIO_LO, COST_RATIO_HI = 0.5, 8.0


def load_prices(which):
    """取对方渠道价：京东 jd_price(opp_mkt)、猫超 channel_price(opp_cost)。"""
    out = {}
    if which == "jd":
        import csv
        v = latest("supply_pull/output/本来生活VOP_goods_本网ID.csv")
        for r in csv.DictReader(v.open(encoding="utf-8-sig")):
            try:
                p = float(r["jd_price"]) if r.get("jd_price") else None
            except ValueError:
                p = None
            if p and p > 0:
                out[str(r["sku_id"])] = p
    else:
        m = latest("supply_pull/output/猫超价格_全量_*.jsonl")
        for l in m.open(encoding="utf-8"):
            if l.strip():
                o = json.loads(l)
                p = o.get("channel_price")
                if isinstance(p, (int, float)) and p > 0:
                    out[str(o["product_id"])] = p
    return out


def cnt_of(name):
    try:
        return (parse_pack(name or "").get("count") or 1) or 1
    except Exception:
        return 1


def build(which, align_file, out_file):
    excl = {str(json.loads(l)["pbasic"]) for l in EXCLUDE.open(encoding="utf-8") if l.strip()}
    names = load_names()
    prices = load_prices(which)
    bl_cost = load_bl_cost()
    dropped = []          # 被合理性护栏丢弃的（留档核查）
    cands = defaultdict(list)
    for l in (HERE / align_file).open(encoding="utf-8"):
        if not l.strip():
            continue
        o = json.loads(l)
        a, b = str(o["a"]), str(o["b"])
        if a in excl or a not in names:
            continue      # 非纯池品 / 无品名（无法折算）跳过
        if b not in prices:
            continue
        if not (LO <= o["sim"] < HI):
            continue
        if o.get("conflict"):
            continue      # 有硬约束冲突不采信
        cnt_bl = cnt_of(names.get(a))
        cnt_opp = cnt_of(o.get("b_name"))
        if cnt_opp <= 0 or cnt_bl <= 0:
            continue
        adj = round(prices[b] / cnt_opp * cnt_bl, 2)
        if adj <= 0:
            continue
        cost = bl_cost.get(a)
        if cost and not (COST_RATIO_LO <= adj / cost <= COST_RATIO_HI):
            dropped.append({"pbasic": a, "target": b, "sim": o["sim"],
                            "price": adj, "cost": cost, "ratio": round(adj / cost, 2),
                            "my_name": names.get(a), "opp_name": o.get("b_name")})
            continue
        cands[a].append({"pbasic": a, "target": b, "sim": o["sim"],
                         "price": adj, "src": f"高边界{sort_name(which)}·规格折算({cnt_bl}/{cnt_opp})"})
    n = 0
    with (HERE / out_file).open("w", encoding="utf-8") as f:
        for a, cs in cands.items():
            best = sorted(cs, key=lambda x: -x["sim"])[0]  # 相似度高优先
            f.write(json.dumps(best, ensure_ascii=False) + "\n")
            n += 1
    # 被丢弃清单留档
    drop_file = out_file.replace("收敛_高边界_", "高边界丢弃_")
    with (HERE / drop_file).open("w", encoding="utf-8") as f:
        for o in dropped:
            f.write(json.dumps(o, ensure_ascii=False) + "\n")
    return n, len(dropped)


def sort_name(which):
    return "京东" if which == "jd" else "猫超"


def main():
    n1, d1 = build("jd", "对齐_向量_bl_jd.jsonl", "收敛_高边界_京东.jsonl")
    print(f"高边界收敛 本网↔京东: {n1} 对（护栏丢弃 {d1}）")
    n2, d2 = build("mc", "对齐_向量_bl_mc.jsonl", "收敛_高边界_猫超.jsonl")
    print(f"高边界收敛 本网↔猫超: {n2} 对（护栏丢弃 {d2}）")


if __name__ == "__main__":
    main()
