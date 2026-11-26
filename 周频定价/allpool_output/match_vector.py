#!/usr/bin/env python3
"""向量同款匹配·全量 pipeline（四货源交叉对齐）。

流程：
  1. 加载四货源品名（本网/京东VOP/猫超SPU/东方甄选），bge-small-zh 编码
  2. FAISS 索引，六向交叉召回 top-K
  3. 硬约束过滤：品牌一致（修好的 brand_tokens）+ 规格归一 + 单价差（有成本/价时）
  4. 分档：≥0.85 自动采信 / 0.70-0.85 边界（待LLM判） / <0.70 丢弃
  5. 产出 对齐_向量_<源>_<目标>.jsonl + 汇总统计

用法：
  HF_HUB_OFFLINE=1 python match_vector.py [--topk 20] [--limit N]
"""
import argparse
import csv
import json
import re
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent  # 周频定价/
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "pricing_output"))
from latest import latest, merged_lines  # noqa: E402
from build_similarity import brand_tokens, parse_pack, brand_main  # noqa: E402

MODEL_NAME = "BAAI/bge-small-zh-v1.5"
ACCEPT_TH = 0.85   # ≥ 自动采信
BORDER_TH = 0.70   # 0.70-0.85 边界待判
TOP_K_DEFAULT = 20


# ---------- 货源加载 ----------

def load_sources(limit=0):
    srcs = {}

    # 本网：pbasic -> {name, price(成本)}
    pool = latest("pricing_output/标准池_四站全上架有库存_*.jsonl")
    info = latest("pricing_output/源数据_标准池商品信息_*.jsonl")
    psys2pb, pb_price = {}, {}
    for l in open(pool, encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            psys2pb[o["productSysNo"]] = o["productBasicSysNo"]
    cost = latest("pricing_output/源数据_标准池ERP站点成本_*.jsonl")
    cost_rows = defaultdict(list)
    for l in open(cost, encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            cost_rows[o["ProductBasicSysNo"]].append(o)
    sys.path.insert(0, str(ROOT / "pricing_output"))
    from build_standard_pool_pricing import resolve_cost
    for pb, rows in cost_rows.items():
        c, _, _ = resolve_cost(rows)
        if c:
            pb_price[pb] = c
    bl = {}
    for l in open(info, encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            pb = psys2pb.get(o["productSysNo"])
            if pb and o.get("name"):
                bl[str(pb)] = {"name": o["name"], "price": pb_price.get(pb)}
    srcs["bl"] = bl

    # 京东VOP：sku_id -> {name, price(sale_price协议价=进货口径), jd_price(C端价另存)}
    jd = {}
    for r in csv.DictReader(open(latest("supply_pull/output/本来生活VOP_goods_本网ID.csv"), encoding="utf-8-sig")):
        def f2(v):
            try:
                return float(v) if v else None
            except ValueError:
                return None
        jd[str(r["sku_id"])] = {"name": r.get("sku_name") or "", "price": f2(r.get("sale_price")),
                                "mkt_price": f2(r.get("jd_price")), "brand": r.get("brand_name") or ""}
    srcs["jd"] = jd

    # 猫超SPU：product_id -> {name, price(供货价SPU=进货口径), mkt_price(券后/一口价另存)}
    mc_price = {}
    pf = latest("supply_pull/output/猫超价格_全量_*.jsonl")
    if pf:
        for l in open(pf, encoding="utf-8"):
            if l.strip():
                o = json.loads(l)
                mc_price[str(o["product_id"])] = o.get("channel_price")
    mc = {}
    for l in merged_lines("supply_pull/output/猫超详情_全量_*.jsonl"):
        d = json.loads(l)["data"]
        pid = str(d.get("product_id"))
        rp = d.get("reserve_price")
        try:
            mkt = round(float(rp) / 100, 2) if rp else None
        except (TypeError, ValueError):
            mkt = None
        mc[pid] = {"name": d.get("product_title") or "", "price": mc_price.get(pid),
                   "mkt_price": mkt, "brand": d.get("brand") or ""}
    srcs["mc"] = mc

    # 东方甄选：specOuterId -> {name, price(purchasePrice采购价=进货口径), mkt_price(retailPrice另存)}
    eb_file = sorted((ROOT / "supply_pull" / "output").glob("东方甄选商品_*.jsonl"))
    eb = {}
    if eb_file:
        for l in open(eb_file[-1], encoding="utf-8"):
            if l.strip():
                o = json.loads(l)
                eb[str(o["specOuterId"])] = {"name": o.get("goodsName") or "",
                                             "price": o.get("purchasePrice"),
                                             "mkt_price": o.get("retailPrice"),
                                             "spec": o.get("specSkuProperties") or ""}
    srcs["eb"] = eb

    if limit:
        srcs = {k: dict(list(v.items())[:limit]) for k, v in srcs.items()}
    return srcs


# ---------- 特征预计算 ----------

# 明确包装件数：数字 + 量词（粒/盒/瓶/袋…）。用于识别"本网有规格、对方无规格"错配。
# 注意：parse_pack 会把"单果克重103-124g"误解析成 base=124g，不可靠，故单独用文本正则。
PACK_COUNT_RE = re.compile(r"(\d+)\s*(粒|个|只|片|袋|盒|瓶|包|桶|支|箱|罐|块|条|张|把|双|组)")


def prep_features(srcs):
    """每个品预计算：品牌token、规格pack、是否含明确包装件数。供硬约束用。"""
    feats = {}
    for sk, items in srcs.items():
        f = {}
        for iid, it in items.items():
            name = it["name"]
            bt = brand_tokens(name)
            bm = brand_main(it.get("brand") or "")
            if bm:
                bt |= {bm[i:i+k] for k in (2, 3, 4) for i in range(max(1, len(bm)-k+1))} | {bm}
            f[iid] = {"brand": bt, "pack": parse_pack(name),
                      "pack_explicit": bool(PACK_COUNT_RE.search(name or ""))}
        feats[sk] = f
    return feats


def hard_ok(fa, fb, pa, pb, enforce_spec=False):
    """硬约束：品牌冲突淘汰；规格都解析出且不一致淘汰；单价差>2.5倍淘汰。

    enforce_spec=True（仅本网侧启用，用户 2026-11-25 拍板）：
      本网品名有明确包装件数（N粒/盒/瓶…），对方既无明确件数、parse_pack 也解析不出规格
      → 判"对方无规格"不采信。典型：佳沛16粒礼盒 vs 猫超无规格散装（语义像但不可比价）。
    """
    ba, bb = fa["brand"], fb["brand"]
    if ba and bb and not (ba & bb):
        return False, "品牌冲突"
    ca, cb = fa["pack"].get("base"), fb["pack"].get("base")
    if ca and cb and ca != cb:
        return False, "规格不一致"
    if enforce_spec and fa.get("pack_explicit") and not cb and not fb.get("pack_explicit"):
        return False, "对方无规格"
    if pa and pb:
        r = pa / pb
        if r > 2.5 or r < 1/2.5:
            return False, "单价差异过大"
    return True, ""


# ---------- 主流程 ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--topk", type=int, default=TOP_K_DEFAULT)
    ap.add_argument("--limit", type=int, default=0, help="每货源只取前N条（试跑）")
    args = ap.parse_args()

    from sentence_transformers import SentenceTransformer

    t0 = time.time()
    print("加载货源…", flush=True)
    srcs = load_sources(args.limit)
    for k, v in srcs.items():
        print(f"  {k}: {len(v)}", flush=True)
    print("预计算品牌/规格特征…", flush=True)
    feats = prep_features(srcs)

    print("加载模型并编码…", flush=True)
    model = SentenceTransformer(MODEL_NAME)
    embs, ids = {}, {}
    for sk, items in srcs.items():
        ids[sk] = list(items.keys())
        texts = [items[i]["name"] for i in ids[sk]]
        embs[sk] = model.encode(texts, batch_size=256, normalize_embeddings=True,
                                show_progress_bar=False).astype("float32")
        print(f"  {sk} 编码 {len(texts)} 条 {time.time()-t0:.0f}s", flush=True)

    # 六向交叉（numpy 分块 top-k：避免整矩阵 75k×145k×4B≈44GB 爆内存；
    # 向量已归一化，内积=余弦；避开 faiss 与 torch 的 OpenMP 冲突）
    pairs = [("bl", "jd"), ("bl", "mc"), ("bl", "eb"),
             ("jd", "mc"), ("jd", "eb"), ("mc", "eb")]
    CHUNK = 2000  # 每块查询行数，峰值内存 ≈ CHUNK×Nb×4B
    summary = {}
    for a, b in pairs:
        if not srcs[a] or not srcs[b]:
            continue
        k = min(args.topk, len(srcs[b]))
        B = embs[b]                       # (Nb, 512)
        stat = Counter()
        out = HERE / f"对齐_向量_{a}_{b}.jsonl"
        with out.open("w", encoding="utf-8") as f:
            Na = len(ids[a])
            for start in range(0, Na, CHUNK):
                end = min(start + CHUNK, Na)
                S = embs[a][start:end] @ B.T          # (chunk, Nb)
                kk = min(k, S.shape[1])
                part = np.argpartition(-S, kth=kk-1, axis=1)[:, :kk]
                rows = np.arange(S.shape[0])[:, None]
                order = np.argsort(-S[rows, part], axis=1)
                I = part[rows, order]
                D = S[rows, I]
                del S
                for ai in range(end - start):
                    aid = ids[a][start + ai]
                    wrote = False
                    for rank in range(kk):
                        bi = int(I[ai][rank])
                        if bi < 0:
                            continue
                        sim = float(D[ai][rank])
                        if sim < BORDER_TH:
                            continue
                        bid = ids[b][bi]
                        ok, why = hard_ok(feats[a][aid], feats[b][bid],
                                          srcs[a][aid]["price"], srcs[b][bid]["price"],
                                          enforce_spec=(a == "bl"))
                        if not ok:
                            stat[f"硬冲突_{why}"] += 1
                        tier = "采信" if (sim >= ACCEPT_TH and ok) else "边界"
                        stat[tier] += 1
                        f.write(json.dumps({"a": aid, "b": bid, "sim": round(sim, 4),
                                            "tier": tier, "rank": rank + 1,
                                            "conflict": None if ok else why,
                                            "a_name": srcs[a][aid]["name"],
                                            "b_name": srcs[b][bid]["name"]},
                                           ensure_ascii=False) + "\n")
                        wrote = True
                    if not wrote:
                        stat["无≥边界候选"] += 1
                if (start // CHUNK) % 10 == 0:
                    print(f"  {a}→{b} {end}/{Na} {time.time()-t0:.0f}s", flush=True)
        summary[f"{a}→{b}"] = dict(stat)
        print(f"{a}→{b}: {dict(stat)}", flush=True)

    print(json.dumps({"耗时s": round(time.time()-t0, 1), "summary": summary},
                     ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
