#!/usr/bin/env python3
"""本网标准池 × 猫超SPU 相似度匹配（score_pair ≥0.75 采信）+ 条码对齐种子合并。

输入：本网 75,067（名称+类目） | 猫超 SPU 144,845（标题+品牌+类目）
输出：allpool_output/对齐_本网猫超.jsonl
      {"pbasic":.., "maochao_pid":.., "score":.., "via":"barcode"|"similarity", ...}
"""
import json
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "pricing_output"))
from build_similarity import score_pair, brand_tokens, brand_main, model_code, pack_of, unit_price, num

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
from latest import latest
POOL = latest("pricing_output/标准池_四站全上架有库存_*.jsonl", "ALLPOOL_POOL")
INFO = latest("pricing_output/源数据_标准池商品信息_*.jsonl", "ALLPOOL_INFO")
MC_DETAIL = latest("supply_pull/output/猫超详情_全量_*.jsonl", "ALLPOOL_MC_DETAIL")
BARCODE_MATCH = HERE / "条码对齐_本网猫超.jsonl"
OUT = HERE / "对齐_本网猫超.jsonl"
THRESHOLD = 0.75
RECORD_FLOOR = 0.60


def load_pool():
    info = {json.loads(l)["productSysNo"]: json.loads(l) for l in INFO.open(encoding="utf-8") if l.strip()}
    items = []
    for l in POOL.open(encoding="utf-8"):
        if not l.strip():
            continue
        o = json.loads(l)
        inf = info.get(o["productSysNo"], {})
        name = inf.get("name") or ""
        items.append({"pbasic": o["productBasicSysNo"], "name": name, "c1": inf.get("c1") or 0})
    for it in items:
        it["_pack"] = pack_of(it["name"])
        it["_brand"] = brand_tokens(it["name"])
        it["_model"] = model_code(it["name"])
    return items


def load_maochao():
    from latest import merged_lines
    items = []
    for l in merged_lines("supply_pull/output/猫超详情_全量_*.jsonl", "ALLPOOL_MC_DETAIL"):
        o = json.loads(l)["data"]
        items.append({"pid": o.get("product_id"), "name": o.get("product_title") or "",
                      "brand": o.get("brand") or "", "category": o.get("category") or ""})
    for it in items:
        it["_pack"] = pack_of(it["name"])
        bm = brand_main(it["brand"])
        bt = brand_tokens(it["name"])
        if bm:
            bt |= {bm[i:i+k] for k in (2, 3, 4) for i in range(max(1, len(bm)-k+1))} | {bm[:2], bm[:3], bm[:4]}
        it["_brand"] = bt
        it["_model"] = model_code(it["name"], bm)
    return items


def main():
    pool = load_pool()
    mc = load_maochao()
    print(f"本网 {len(pool)} × 猫超SPU {len(mc)}", flush=True)

    mc_by_brand = defaultdict(list)
    mc_no_brand = []
    for v in mc:
        if v["_brand"]:
            for t in v["_brand"]:
                mc_by_brand[t].append(v)
        else:
            mc_no_brand.append(v)

    barcode_hit = {}
    for l in BARCODE_MATCH.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            barcode_hit.setdefault(o["pbasic"], o["maochao_pid"])

    n = matched = 0
    best_low = 0
    with OUT.open("w", encoding="utf-8") as f:
        # 先写条码对齐种子
        for pb, pid in barcode_hit.items():
            f.write(json.dumps({"pbasic": pb, "maochao_pid": pid, "score": 1.0,
                                "via": "barcode"}, ensure_ascii=False) + "\n")
        for i, p in enumerate(pool):
            n += 1
            if p["pbasic"] in barcode_hit:
                matched += 1
                continue
            cand, seen = [], set()
            for t in p["_brand"]:
                for v in mc_by_brand.get(t, []):
                    if v["pid"] not in seen:
                        seen.add(v["pid"])
                        cand.append(v)
            best = None
            for v in cand:
                our_unit = unit_price(None, p["_pack"])
                mc_unit = unit_price(None, v["_pack"])
                total, bf, mf, sf, pf, catf, *_ = score_pair(
                    p["name"], v["name"], vop_brand=v["brand"],
                    lp_model=p["_model"], vop_model=v["_model"],
                    lp_pack=p["_pack"], vop_pack=v["_pack"],
                    our_price=our_unit, jd_price=mc_unit)
                if best is None or total > best[0]:
                    best = (total, v)
            if best and best[0] >= THRESHOLD:
                matched += 1
                f.write(json.dumps({"pbasic": p["pbasic"], "maochao_pid": best[1]["pid"],
                                    "score": round(best[0], 4), "via": "similarity",
                                    "pool_name": p["name"], "mc_name": best[1]["name"]},
                                   ensure_ascii=False) + "\n")
            elif best and best[0] >= RECORD_FLOOR:
                best_low += 1
            if (i + 1) % 5000 == 0:
                print(f"{i+1}/{len(pool)}: 命中 {matched} ({matched/(i+1):.1%})", flush=True)

    print(json.dumps({"本网品数": n, "对齐总数(条码+相似度)": matched,
                      "对齐率": f"{matched/n:.1%}", "条码种子": len(barcode_hit),
                      "0.6-0.75未采信": best_low}, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
