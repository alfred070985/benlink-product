#!/usr/bin/env python3
"""标准池商品 × 京东VOP 相似品自动匹配：score_pair ≥0.75 直接采信为同品。

复用 build_similarity.py 的评分函数。品牌预过滤 + 一级类目(c1)分组内比对，控制计算量。
输入：源数据_标准池商品信息_<日期>.jsonl + supply_pull/output/本来生活VOP_goods_本网ID.csv
输出：标准池京东相似匹配_<日期>.jsonl（每行：productSysNo, jd_sku, score, jd_price, 匹配名）
"""
import csv
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_similarity import (
    score_pair, brand_tokens, brand_main, model_code, pack_of, unit_price, num,
)

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")
POOL_INFO = PRICING_DIR / f"源数据_标准池商品信息_{TODAY}.jsonl"
VOP_CSV = HERE.parent / "supply_pull" / "output" / "本来生活VOP_goods_本网ID.csv"
OUT = PRICING_DIR / f"标准池京东相似匹配_{TODAY}.jsonl"
THRESHOLD = 0.75
RECORD_FLOOR = 0.60   # 0.6 以上全部记录，供阈值决策


def load_pool(limit=None):
    items = []
    with POOL_INFO.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                items.append(json.loads(line))
    if limit:
        items = items[:limit]
    for it in items:
        name = it.get("name") or ""
        it["_pack"] = pack_of(name)
        it["_brand"] = brand_tokens(name)
        it["_model"] = model_code(name)
        it["_c1"] = it.get("c1") or 0
    return items


def load_vop():
    items = []
    with VOP_CSV.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            jd = num(r.get("jd_price"))
            if not jd:
                continue
            name = r.get("sku_name") or ""
            items.append({
                "sku": r.get("sku_id"),
                "name": name,
                "brand_raw": r.get("brand_name") or "",
                "jd_price": jd,
                "c": r.get("category") or "",
            })
    for it in items:
        it["_pack"] = pack_of(it["name"])
        bm = brand_main(it["brand_raw"])
        bt = brand_tokens(it["name"])
        if bm:
            bt |= {bm[i:i+k] for k in (2, 3, 4) for i in range(max(1, len(bm)-k+1))} | {bm[:2], bm[:3], bm[:4]}
        it["_brand"] = bt
        it["_model"] = model_code(it["name"], bm)
    return items


def brand_maybe(lp, vop):
    """品牌预过滤：任一方无品牌则放行（无法判断），有品牌则要求交集。"""
    if not lp or not vop:
        return True
    return bool(lp & vop)


def match_one(pool_item, vop_candidates):
    best = None
    for v in vop_candidates:
        if not brand_maybe(pool_item["_brand"], v["_brand"]):
            continue
        our_unit = unit_price(pool_item.get("ref_price"), pool_item["_pack"])
        jd_unit = unit_price(v["jd_price"], v["_pack"])
        total, bf, mf, sf, pf, catf, *_ = score_pair(
            pool_item["name"], v["name"], vop_brand=v["brand_raw"],
            lp_model=pool_item["_model"], vop_model=v["_model"],
            lp_pack=pool_item["_pack"], vop_pack=v["_pack"],
            our_price=our_unit, jd_price=jd_unit,
        )
        if best is None or total > best[0]:
            best = (total, v, (bf, mf, sf, pf, catf))
    return best


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else None
    pool = load_pool(limit)
    vop = load_vop()
    print(f"标准池 {len(pool)} × 京东VOP {len(vop)}", flush=True)

    # 京东侧按 c1 无法对齐（京东类目编码不同），按品牌倒排索引加速：
    # 对每个标准池商品，只比对"品牌token有交集 或 京东无品牌token"的候选
    vop_by_brand = defaultdict(list)
    vop_no_brand = []
    for v in vop:
        if v["_brand"]:
            for t in v["_brand"]:
                vop_by_brand[t].append(v)
        else:
            vop_no_brand.append(v)

    matched = 0
    scores = []
    with OUT.open("w", encoding="utf-8") as f:
        for i, p in enumerate(pool):
            # 候选：品牌token命中的 + 无品牌的（保守起见只取品牌桶，无品牌京东品太多会拖慢，先跳过）
            cand = []
            seen = set()
            for t in p["_brand"]:
                for v in vop_by_brand.get(t, []):
                    if v["sku"] not in seen:
                        seen.add(v["sku"])
                        cand.append(v)
            best = match_one(p, cand) if cand else None
            if best and best[0] >= RECORD_FLOOR:
                total, v, flags = best
                if total >= THRESHOLD:
                    matched += 1
                    scores.append(total)
                f.write(json.dumps({
                    "productSysNo": p["productSysNo"],
                    "jd_sku": v["sku"],
                    "score": round(total, 4),
                    "采纳": total >= THRESHOLD,
                    "jd_price": v["jd_price"],
                    "jd_name": v["name"],
                    "pool_name": p["name"],
                    "flags": flags,
                }, ensure_ascii=False) + "\n")
            if (i + 1) % 1000 == 0:
                print(f"{i+1}/{len(pool)}: 命中 {matched} ({matched/(i+1):.1%})", flush=True)

    import statistics
    print(json.dumps({
        "处理商品数": len(pool), "采信数(>=0.75)": matched,
        "采信率": round(matched / len(pool), 4) if pool else 0,
        "分数中位数": round(statistics.median(scores), 4) if scores else None,
        "阈值": THRESHOLD,
        "输出": str(OUT),
    }, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
