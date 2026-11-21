#!/usr/bin/env python3
"""京东 VOP ↔ (本网 + 猫超SPU) 相似度匹配：score_pair ≥0.75 采信。

京东侧：本来生活VOP_goods_本网ID.csv（44,387 个有 jd_price 的 SKU）
候选池：本网 75,067（pbasic）+ 猫超 SPU 144,845（pid），品牌倒排预过滤
输出：allpool_output/对齐_京东.jsonl
      {"jd_sku":.., "target_id":..(pbasic或maochao_pid), "target_src":"bl"|"maochao", "score":..}
"""
import csv
import json
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "pricing_output"))
from build_similarity import score_pair, brand_tokens, brand_main, model_code, pack_of, unit_price, num

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
VOP = ROOT / "jdvop_pull" / "output" / "本来生活VOP_goods_本网ID.csv"
INFO = ROOT / "pricing_output" / "源数据_标准池商品信息_2026-11-19.jsonl"
POOL = ROOT / "pricing_output" / "标准池_四站全上架有库存_2026-11-19.jsonl"
MC_DETAIL = ROOT / "jdvop_pull" / "output" / "猫超详情_全量_2026-11-20.jsonl"
OUT = HERE / "对齐_京东.jsonl"
THRESHOLD = 0.75


def load_targets(which):
    """候选池：bl=本网 75,067 / mc=猫超SPU 144,845，分开跑降低内存峰值"""
    info = {json.loads(l)["productSysNo"]: json.loads(l) for l in INFO.open(encoding="utf-8") if l.strip()}
    targets = []
    if which == "bl":
        for l in POOL.open(encoding="utf-8"):
            if not l.strip():
                continue
            o = json.loads(l)
            inf = info.get(o["productSysNo"], {})
            name = inf.get("name") or ""
            targets.append({"tid": f"bl:{o['productBasicSysNo']}", "src": "bl", "name": name, "brand": ""})
    else:
        for l in MC_DETAIL.open(encoding="utf-8"):
            if not l.strip():
                continue
            d = json.loads(l)["data"]
            targets.append({"tid": f"mc:{d.get('product_id')}", "src": "maochao",
                            "name": d.get("product_title") or "", "brand": d.get("brand") or ""})
    for t in targets:
        t["pack"] = pack_of(t["name"])
        bt = brand_tokens(t["name"])
        bm = brand_main(t["brand"])
        if bm:
            bt |= {bm[i:i+k] for k in (2, 3, 4) for i in range(max(1, len(bm)-k+1))} | {bm[:2], bm[:3], bm[:4]}
        t["brand_toks"] = bt
        t["model"] = model_code(t["name"], bm)
    return targets


def main():
    which = sys.argv[1] if len(sys.argv) > 1 else "bl"
    targets = load_targets(which)
    by_brand = defaultdict(list)
    for t in targets:
        for tok in t["brand_toks"]:
            by_brand[tok].append(t)
    del targets
    print(f"候选池={which}，词条 {len(by_brand)}", flush=True)

    n = matched = low = 0
    out_path = OUT.with_name(f"对齐_京东_{which}.jsonl")
    with out_path.open("w", encoding="utf-8") as f:
        with VOP.open(encoding="utf-8-sig", newline="") as vf:
            for r in csv.DictReader(vf):
                jd_price = num(r.get("jd_price"))
                if not jd_price:
                    continue
                n += 1
                name = r.get("sku_name") or ""
                bm = brand_main(r.get("brand_name"))
                bt = brand_tokens(name)
                if bm:
                    bt |= {bm[i:i+k] for k in (2, 3, 4) for i in range(max(1, len(bm)-k+1))} | {bm[:2], bm[:3], bm[:4]}
                jp = pack_of(name)
                jm = model_code(name, bm)
                ju = unit_price(jd_price, jp)

                cand, seen = [], set()
                for t in bt:
                    for v in by_brand.get(t, []):
                        if v["tid"] not in seen:
                            seen.add(v["tid"])
                            cand.append(v)
                best = None
                for v in cand:
                    total, *_, = score_pair(
                        name, v["name"], vop_brand=r.get("brand_name") or "",
                        lp_model=jm, vop_model=v["model"],
                        lp_pack=jp, vop_pack=v["pack"],
                        our_price=ju, jd_price=unit_price(None, v["pack"]))
                    if best is None or total > best[0]:
                        best = (total, v)
                if best and best[0] >= THRESHOLD:
                    matched += 1
                    t = best[1]
                    tid_val = t["tid"].split(":", 1)[1]
                    f.write(json.dumps({"jd_sku": r.get("sku_id"), "jd_name": name[:80],
                                        "target_id": tid_val, "target_src": t["src"],
                                        "score": round(best[0], 4)}, ensure_ascii=False) + "\n")
                elif best and best[0] >= 0.6:
                    low += 1
                if n % 5000 == 0:
                    print(f"{n}: 命中 {matched} ({matched/n:.1%})", flush=True)

    print(json.dumps({"京东SKU数": n, "命中": matched, "命中率": f"{matched/n:.1%}",
                      "0.6-0.75未采信": low}, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
