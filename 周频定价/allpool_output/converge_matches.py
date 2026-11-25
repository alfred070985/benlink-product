#!/usr/bin/env python3
"""向量候选收敛为 1对1 匹配表（本网品 ↔ 京东SKU / 猫超SPU）。

收敛规则（用户确认方向）：同一本网品的多条候选中——
  1. 硬约束无冲突优先（规格一致/无冲突 > 有冲突）
  2. 进货价最接近优先（|本方价-对方价|/本方价 最小）
  3. 相似度最高兜底
只处理采信档（≥0.85）；边界档留待 LLM 裁决层。

输出：
  收敛_本网京东.jsonl   {pbasic, sku_id, sim, jd_price, sale_price}
  收敛_本网猫超.jsonl   {pbasic, product_id, sim, coupon_price, reserve_price}
"""
import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from latest import latest  # noqa: E402


def load_prices():
    """本方(本网成本)与对方进货价，用于价格接近度。"""
    pool = latest("pricing_output/标准池_四站全上架有库存_*.jsonl")
    from build_standard_pool_pricing import resolve_cost
    cost_rows = defaultdict(list)
    for l in open(latest("pricing_output/源数据_标准池ERP站点成本_*.jsonl"), encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            cost_rows[o["ProductBasicSysNo"]].append(o)
    bl_price = {}
    for pb, rows in cost_rows.items():
        c, _, _ = resolve_cost(rows)
        if c:
            bl_price[pb] = c
    import csv
    jd_price = {}
    for r in csv.DictReader(open(latest("supply_pull/output/本来生活VOP_goods_本网ID.csv"), encoding="utf-8-sig")):
        def f(v):
            try:
                return float(v) if v else None
            except ValueError:
                return None
        jd_price[str(r["sku_id"])] = (f(r.get("sale_price")), f(r.get("jd_price")))
    mc_price = {}
    for l in open(latest("supply_pull/output/猫超价格_全量_*.jsonl"), encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            mc_price[str(o["product_id"])] = (o.get("channel_price"), o.get("coupon_price"))
    return bl_price, jd_price, mc_price


def converge(align_file, out_file, my_price, opp_price_map, opp_price_idx=0):
    """按 (冲突少→价格近→相似度高) 收敛，每个本方品取一条。"""
    cands = defaultdict(list)
    for l in open(HERE / align_file, encoding="utf-8"):
        o = json.loads(l)
        if o["tier"] != "采信":
            continue
        cands[o["a"]].append(o)
    n = 0
    with open(HERE / out_file, "w", encoding="utf-8") as f:
        for a, cands_a in cands.items():
            p0 = my_price.get(a) or my_price.get(int(a) if str(a).isdigit() else a)
            def key(o):
                conflict = 1 if o.get("conflict") else 0
                pp = opp_price_map.get(o["b"])
                op = pp[opp_price_idx] if pp else None
                pdiff = abs(op - p0) / p0 if (p0 and op) else 9.9
                return (conflict, pdiff, -o["sim"])
            best = sorted(cands_a, key=key)[0]
            pp = opp_price_map.get(best["b"]) or (None, None)
            f.write(json.dumps({"pbasic": a, "target": best["b"], "sim": best["sim"],
                                "conflict": best.get("conflict"),
                                "opp_cost": pp[0], "opp_mkt": pp[1]},
                               ensure_ascii=False) + "\n")
            n += 1
    return n


def main():
    bl_price, jd_price, mc_price = load_prices()
    n1 = converge("对齐_向量_bl_jd.jsonl", "收敛_本网京东.jsonl", bl_price, jd_price, 0)
    print(f"收敛 本网↔京东: {n1} 对")
    n2 = converge("对齐_向量_bl_mc.jsonl", "收敛_本网猫超.jsonl", bl_price, mc_price, 0)
    print(f"收敛 本网↔猫超: {n2} 对")


if __name__ == "__main__":
    sys.path.insert(0, str(HERE.parent / "pricing_output"))
    main()
