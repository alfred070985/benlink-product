#!/usr/bin/env python3
"""货盘一重建 v2：纯本来池 + 向量升级后的市场价 + 空间率重算。

相对 build_standard_pool_pricing.py 的三处变更（其余口径全部沿用）：
  1. 剔除 ERP t_goods_map 标记的 16,644 京东品 → 纯本来池 58,423
  2. 对标价新增「向量收敛新价」来源（收敛_本网京东/猫超.jsonl，1对1）
  3. 优先级重排：天镜人工核实 > 向量收敛新价 > 旧算法相似采信 > ERP兜底

保持不变（用户拍板业务红线）：
  - 成本口径：resolve_cost（展示优先→未展示回退、排除作废；直发采购价/仓发含税成本）
  - 加价档：直发 2/5/8%、仓发 5/12/24%
  - 选档规则：默认中档②；档②空间<10%降①；≥30%且档③≥10%升③；兜底价不升档
  - 售价字段：京东 jd_price、猫超 channel_price（与旧口径一致的渠道/协议价）

只读源文件，输出全新工作簿。路径由 __file__ 推导，latest() 解析跨周文件。
"""
import json
import os
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
ALLOUT = ROOT / "allpool_output"
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ALLOUT))

from latest import latest  # noqa: E402
import build_standard_pool_pricing as base  # noqa: E402  复用成本/档位/选档口径

# base 模块按"今天"拼源文件名，但源文件在周归档里 → 用 latest() 解析后覆盖其路径常量
base.POOL = latest("pricing_output/标准池_四站全上架有库存_*.jsonl", "PRICING_POOL")
base.INFO = latest("pricing_output/源数据_标准池商品信息_*.jsonl", "PRICING_INFO")
base.COST = latest("pricing_output/源数据_标准池ERP站点成本_*.jsonl", "PRICING_COST")
base.MATCH = latest("pricing_output/标准池京东相似匹配_*.jsonl", "PRICING_MATCH")
base.HIST_JSONL = latest("pricing_output/源数据_天镜价格历史全量_*.jsonl")
base.MAP_JSONL = latest("pricing_output/源数据_天镜同品映射_*.jsonl")

TODAY = datetime.now().strftime("%Y-%m-%d")
OUT = HERE / f"分销标准池定价_{TODAY}.xlsx"
STATS_OUT = HERE / f"定价统计_{TODAY}.json"
# 跨周底表：latest() 全仓解析（归档后工作目录无文件也能重跑）
EXCLUDE = latest("allpool_output/分销池_ERP京东品.jsonl", "ERP_JD_EXCLUDE")
CONV_JD = latest("allpool_output/收敛_本网京东.jsonl", "CONV_JD")
CONV_MC = latest("allpool_output/收敛_本网猫超.jsonl", "CONV_MC")
# 高边界（0.80-0.85 无冲突·规格折算）补充来源，用户 2026-11-25 拍板降档采信
HB_JD = latest("allpool_output/收敛_高边界_京东.jsonl", "HB_JD")
HB_MC = latest("allpool_output/收敛_高边界_猫超.jsonl", "HB_MC")


def load_vector_price():
    """向量收敛新价：{pbasic: (market_price, source, sku)}。
    京东取 jd_price(opp_mkt)，猫超取 channel_price(opp_cost)，与旧口径一致。
    同一品京东/猫超都有时：优先京东（原对标主源），猫超补缺。"""
    out = {}
    hb_mc_ok = HB_MC is not None and HB_MC.exists()
    hb_jd_ok = HB_JD is not None and HB_JD.exists()
    # 优先级从低到高填：高边界 < 正式收敛；同类京东覆盖猫超
    # 1) 高边界·猫超（折算价）
    for l in (HB_MC.open(encoding="utf-8") if hb_mc_ok else []):
        if l.strip():
            o = json.loads(l)
            if o.get("price") and o["price"] > 0:
                out[str(o["pbasic"])] = (o["price"], o.get("src", "高边界·猫超折算"), str(o["target"]))
    # 2) 高边界·京东（覆盖猫超）
    for l in (HB_JD.open(encoding="utf-8") if hb_jd_ok else []):
        if l.strip():
            o = json.loads(l)
            if o.get("price") and o["price"] > 0:
                out[str(o["pbasic"])] = (o["price"], o.get("src", "高边界·京东折算"), str(o["target"]))
    # 3) 正式收敛·猫超（≥0.85，覆盖高边界）
    for l in CONV_MC.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            p = o.get("opp_cost")
            if p and p > 0:
                out[str(o["pbasic"])] = (p, "向量收敛·猫超渠道价", str(o["target"]))
    # 4) 正式收敛·京东（≥0.85，最高优先）
    for l in CONV_JD.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            p = o.get("opp_mkt")
            if p and p > 0:
                out[str(o["pbasic"])] = (p, "向量收敛·京东渠道价", str(o["target"]))
    return out


def main():
    # 加载剔除清单
    excl = {str(json.loads(l)["pbasic"]) for l in EXCLUDE.open(encoding="utf-8") if l.strip()}
    print(f"ERP京东品剔除清单: {len(excl)}", flush=True)

    # 复用原脚本的数据加载（POOL/INFO/COST/MATCH/HIST/MAP 均由环境变量或 __file__ 推导）
    pool = {json.loads(l)["productSysNo"]: json.loads(l) for l in base.POOL.open(encoding="utf-8") if l.strip()}
    info = {json.loads(l)["productSysNo"]: json.loads(l) for l in base.INFO.open(encoding="utf-8") if l.strip()}
    costs = base.load_costs()
    hist, same, diff, sim = base.load_tianjing()
    match = {}
    for l in base.MATCH.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            if o.get("采纳"):
                match[o["productSysNo"]] = o
    vec = load_vector_price()
    print(f"标准池 {len(pool)} | 采信对标(旧) {len(match)} | 向量新价 {len(vec)}", flush=True)

    records = []
    stat = Counter()
    for psys in sorted(pool):
        p = pool[psys]
        pb = p["productBasicSysNo"]
        if str(pb) in excl:          # ← 核心：从源头剔除京东品
            stat["剔除_京东品"] += 1
            continue
        inf = info.get(psys, {})
        cost, basis, direct = base.resolve_cost(costs.get(pb, []))
        m = match.get(psys)

        jd = jd_src = jd_sku = None
        # 优先级1：天镜人工核实·同款 → 天镜价，无价回退VOP
        skus = same.get(pb) or []
        tj = next((hist[s] for s in skus if s in hist), None)
        if tj is not None:
            jd, jd_src, jd_sku = tj, "天镜智价·人工核实同款", skus[0]
            stat["对标_天镜同款"] += 1
        elif skus and m:
            jd, jd_src, jd_sku = m["jd_price"], "京东VOP·同款映射备选价", m["jd_sku"]
            stat["对标_VOP备选"] += 1
        # 优先级2：天镜不同规 → 折算
        elif diff.get(pb):
            vs, ratio = diff[pb][0]
            tj = hist.get(vs)
            if tj is not None:
                from build_similarity import parse_pack
                try:
                    a, b = (int(x) for x in str(ratio).split(":"))
                    if a <= 0 or b <= 0:
                        raise ValueError
                except Exception:
                    a, b = 1, 1
                cnt = parse_pack(inf.get("name") or "").get("count") or 1
                jd = round(tj / a * cnt, 2)
                jd_src = f"天镜智价·不同规折算({ratio}·{cnt}件)"
                jd_sku = vs
                stat["对标_天镜不同规"] += 1
            else:
                stat["待确认_不同规无价"] += 1
        # 优先级3：向量收敛新价（用户拍板优先级：天镜 > 向量新价 > 旧算法）
        elif str(pb) in vec:
            jd, jd_src, jd_sku = vec[str(pb)]
            stat["对标_高边界折算" if "高边界" in jd_src else "对标_向量新价"] += 1
        # 优先级4：旧算法相似采信(≥0.75)
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
                stat["对标_ERP兜底"] += 1
            else:
                jd, jd_src, jd_sku = None, "无对标价", ""
                stat["对标_无"] += 1
        if cost is None:
            stat["成本缺失"] += 1
        tiers = base.price_tiers(cost, direct)
        idx, reason = base.suggest(tiers, jd, jd_src.startswith("ERP基准价"))
        records.append({
            "psys": psys, "pbasic": pb,
            "name": inf.get("name", ""), "c1": inf.get("c1"), "c2": inf.get("c2"), "c3": inf.get("c3"),
            "ship": "直发" if direct else "仓发", "cost": cost, "basis": basis,
            "jd": jd, "jd_src": jd_src, "jd_sku": jd_sku,
            "tiers": tiers, "suggest_idx": idx, "suggest_reason": reason,
        })
        stat["直发" if direct else "仓发"] += 1

    # 直接复用原脚本的 build_workbook（保持成品结构完全一致）
    base.build_workbook.__globals__["OUT"] = OUT
    base.build_workbook(records, stat)

    # 统计落盘
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
            spaces[key].append((r["jd"] - r["tiers"][r["suggest_idx"]][1]) / r["jd"])
        if r["cost"] is not None:
            for t in range(3):
                chk_cells += 1
                if abs(r["tiers"][t][1] - round(r["cost"] * (1 + ratios_of[r["ship"]][t]), 2)) > 0.011:
                    chk_err += 1
    stats = {
        "日期": TODAY, "版本": "v2_纯池+向量新价",
        "总行数": len(records), "剔除_京东品": stat.get("剔除_京东品", 0),
        "发货类型": dict(ship_stat),
        "对标来源": {k.replace("对标_", ""): v for k, v in stat.items() if k.startswith("对标_")},
        "建议档分布": {k: dict(v) for k, v in suggest_stat.items()},
        "建议档空间率中位数": {
            "真实对标": round(statistics.median(spaces["真实对标"]), 4) if spaces["真实对标"] else None,
            "ERP兜底": round(statistics.median(spaces["ERP兜底"]), 4) if spaces["ERP兜底"] else None,
        },
        "公式校验": {"格数": chk_cells, "错误数": chk_err},
    }
    STATS_OUT.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
