#!/usr/bin/env python3
"""向量同款匹配·验证脚本（先验证后全量）。

用两套独立"标准答案"测 bge-small-zh 向量召回质量：
  A. 天镜人工核实映射（本网P码 ↔ 京东SKU，same_sku 类型）
  B. 猫超69码精确对（本网 ↔ 猫超，条码级 ground truth）

指标：
  - 召回率：真对在本网→对方 top-K 中的命中率（K=1/5/20）
  - 分布：真对相似度 vs 随机负样本相似度（看可分性）

用法：python validate_embedding.py
"""
import json
import random
import sys
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
sys.path.insert(0, str(HERE))
from latest import latest  # noqa: E402

MODEL_NAME = "BAAI/bge-small-zh-v1.5"


def load_bl_names():
    """本网 pbasic -> 名称"""
    info = latest("pricing_output/源数据_标准池商品信息_*.jsonl")
    pool = latest("pricing_output/标准池_四站全上架有库存_*.jsonl")
    psys2pb = {}
    for l in open(pool, encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            psys2pb[o["productSysNo"]] = o["productBasicSysNo"]
    out = {}
    for l in open(info, encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            pb = psys2pb.get(o["productSysNo"])
            if pb and o.get("name"):
                out[pb] = o["name"]
    return out


def load_jd_names():
    """京东 sku_id -> 名称"""
    import csv
    v = latest("supply_pull/output/本来生活VOP_goods_本网ID.csv")
    out = {}
    for r in csv.DictReader(open(v, encoding="utf-8-sig")):
        out[str(r["sku_id"])] = r.get("sku_name") or ""
    return out


def load_mc_names():
    """猫超 product_id -> 标题（跨历史详情文件并集）"""
    from latest import merged_lines
    out = {}
    for l in merged_lines("supply_pull/output/猫超详情_全量_*.jsonl"):
        d = json.loads(l)["data"]
        out[str(d.get("product_id"))] = d.get("product_title") or ""
    return out


def load_tianjing_pairs():
    """天镜人工核实 same_sku: [(pbasic, jd_sku)]"""
    m = latest("pricing_output/源数据_天镜同品映射_*.jsonl")
    pairs = []
    for l in open(m, encoding="utf-8"):
        if not l.strip():
            continue
        o = json.loads(l)
        ic = str(o.get("internal_code") or "").strip()
        if o.get("relation_type") == "same_sku" and ic and ic[0].upper() == "P" and ic[1:].isdigit():
            pairs.append((int(ic[1:]), str(o.get("vop_sku") or "").strip()))
    return pairs


def load_69_pairs():
    """69码条码对: [(pbasic, 猫超product_id)]（SPU级）"""
    f = HERE / "对齐_猫超69码.jsonl"
    pairs = []
    for l in open(f, encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            if o["level"] == "spu":
                pairs.append((o["pbasic"], None))  # product_id 需从条码索引反查
    # 反查 product_id
    idx = latest("猫超条码索引.jsonl") or HERE / "猫超条码索引.jsonl"
    bc2pid = {}
    for l in open(idx, encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            if not o.get("sku_id"):
                bc2pid[o["barcode"]] = str(o["product_id"])
    out = []
    for l in open(f, encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            if o["level"] == "spu" and o["barcode"] in bc2pid:
                out.append((o["pbasic"], bc2pid[o["barcode"]]))
    return out


def encode(model, texts, bs=64):
    return model.encode(texts, batch_size=bs, normalize_embeddings=True, show_progress_bar=False)


def evaluate(name, pairs, src_names, tgt_names, model):
    # 只保留两边都有名字的
    valid = [(s, t) for s, t in pairs if s in src_names and t in tgt_names]
    print(f"\n=== {name}: 有效真对 {len(valid)}/{len(pairs)} ===")
    if not valid:
        return
    tgt_ids = list(tgt_names.keys())
    tgt_texts = [tgt_names[i] for i in tgt_ids]
    tgt_emb = encode(model, tgt_texts)
    # 真对相似度
    sims, hits = [], {1: 0, 5: 0, 20: 0}
    for s, t in valid:
        q = encode(model, [src_names[s]])[0]
        sim = tgt_emb @ q
        rank = int((sim > sim[tgt_ids.index(t)]).sum()) + 1
        sims.append(float(sim[tgt_ids.index(t)]))
        for k in hits:
            if rank <= k:
                hits[k] += 1
    # 随机负样本
    random.seed(42)
    negs = []
    for _ in range(2000):
        s = random.choice(list(src_names.values()))
        t = random.choice(tgt_texts)
        negs.append((s, t))
    qs = encode(model, [s for s, _ in negs])
    ts = encode(model, [t for _, t in negs])
    neg_sims = (qs * ts).sum(axis=1)
    sims = np.array(sims)
    print(f"真对相似度: 中位 {np.median(sims):.3f}  P25 {np.percentile(sims,25):.3f}  P75 {np.percentile(sims,75):.3f}")
    print(f"负样本相似度: 中位 {np.median(neg_sims):.3f}  P95 {np.percentile(neg_sims,95):.3f}  P99 {np.percentile(neg_sims,99):.3f}")
    print(f"召回率: top1 {hits[1]/len(valid):.1%}  top5 {hits[5]/len(valid):.1%}  top20 {hits[20]/len(valid):.1%}")
    for th in (0.75, 0.80, 0.85, 0.90):
        print(f"  阈值{th}: 真对过线 {(sims>=th).mean():.1%} | 负样本过线 {(neg_sims>=th).mean():.2%}")


def main():
    from sentence_transformers import SentenceTransformer
    print(f"加载模型 {MODEL_NAME} …", flush=True)
    model = SentenceTransformer(MODEL_NAME)
    bl = load_bl_names()
    jd = load_jd_names()
    print(f"本网 {len(bl)} 品名 | 京东 {len(jd)} 品名", flush=True)
    evaluate("A. 天镜人工核实（本网↔京东）", load_tianjing_pairs(), bl, jd, model)
    mc = load_mc_names()
    print(f"猫超 {len(mc)} 品名", flush=True)
    evaluate("B. 69码条码对（本网↔猫超）", load_69_pairs(), bl, mc, model)


if __name__ == "__main__":
    main()
