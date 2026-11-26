#!/usr/bin/env python3
"""商品同款匹配 v2 流水线（行业成熟方案四层中的②③④）。

用户 2026-11-25 拍板换方案：弃用"纯向量余弦定采信"，改行业成熟流水线：
  ① 属性抽取 + 强约束（attr_extract.py，已在上游完成）
  ② 多路召回（向量 top-N 候选池，仅作召回不当裁判）
  ③ Cross-Encoder 精排（bge-reranker-v2-m3，逐对打分做裁判）
  ④ 图聚类 + 真条码验证（2,539 对 69码真同款校准阈值）

输入：对齐_向量_bl_jd.jsonl / 对齐_向量_bl_mc.jsonl（召回候选池）
输出：对齐_v2_bl_jd.jsonl / 对齐_v2_bl_mc.jsonl（带 rerank_score、属性冲突、tier）
      v2_阈值校准.json（真条码验证集上的阈值-召回率曲线）

用法：
  HF_HUB_OFFLINE=1 python match_v2_pipeline.py --topn 5 [--limit N]
"""
import argparse
import csv
import json
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "pricing_output"))

from latest import latest  # noqa: E402
from build_similarity import brand_tokens  # noqa: E402
from attr_extract import attrs_of  # noqa: E402

EXCLUDE = ROOT / "交付物/2026-W48/03_货盘底表/5_跨源对齐/分销池_ERP京东品.jsonl"
ALIGN_DIR = HERE
MODEL = "BAAI/bge-reranker-v2-m3"


def load_exclude():
    return {str(json.loads(l)["pbasic"]) for l in EXCLUDE.open(encoding="utf-8") if l.strip()}


def load_recall(fn, topn, excl, limit=0):
    """召回层：每本网品取向量分最高的 topn 个候选（不管旧 tier）。"""
    cands = defaultdict(list)
    n = 0
    for l in (ALIGN_DIR / fn).open(encoding="utf-8"):
        if not l.strip():
            continue
        o = json.loads(l)
        a = str(o["a"])
        if a in excl:
            continue
        if limit and n >= limit:
            break
        cands[a].append(o)
        n += 1
    out = {}
    for a, cs in cands.items():
        cs.sort(key=lambda x: -x["sim"])
        out[a] = cs[:topn]
    return out


def attr_filter(recall):
    """属性强约束：品牌冲突/件数不等/规格不等 → 直接否；返回保留对 + 统计。"""
    kept = []
    stat = Counter()
    for a, cs in recall.items():
        for o in cs:
            aa, bb = attrs_of(o.get("a_name")), attrs_of(o.get("b_name"))
            ba, bb_ = brand_tokens(o.get("a_name")), brand_tokens(o.get("b_name"))
            why = ""
            if ba and bb_ and not (ba & bb_):
                why = "品牌冲突"
            elif aa["pack_explicit"] and bb["pack_explicit"] and aa["count"] != bb["count"]:
                why = "件数不等"
            elif aa["base"] and bb["base"] and aa["base"] != bb["base"]:
                why = "规格不等"
            elif aa["pack_explicit"] and not bb["pack_explicit"] and not bb["base"]:
                # 用户 11-25 拍板延续：本网有明确件数、对方既无件数也无规格 → 不可比价，拦
                # 典型：佳沛16粒礼盒 vs 猫超无规格散装（向量0.851/精排0.956 均误判）
                why = "对方无规格"
            if why:
                stat[why] += 1
            else:
                kept.append((o, aa, bb))
                stat["保留"] += 1
    return kept, stat


def load_truth():
    """69码真同款验证集：{pbasic: [(mc_product_id, bc),...]}"""
    g = [json.loads(l) for l in (HERE / "对齐_猫超69码.jsonl").open(encoding="utf-8") if l.strip()]
    mc = latest("supply_pull/output/猫超商品宽表SKU级_*.csv")
    bar2 = {}
    for r in csv.DictReader(open(mc, encoding="utf-8-sig")):
        bc = (r.get("sku_barcode") or "").strip()
        if bc and bc not in bar2:
            bar2[bc] = (str(r.get("product_id")), r.get("product_title") or "")
    truth = defaultdict(list)
    for x in g:
        if x["barcode"] in bar2:
            pid, nm = bar2[x["barcode"]]
            truth[str(x["pbasic"])].append((pid, nm))
    return truth


def calibrate(pairs, scores, truth, truth_scores):
    """用真条码对校准阈值：输出 (阈值, 真同款通过率, 全量保留率) 曲线。"""
    import numpy as np
    all_s = np.array(scores)
    true_s = np.array(truth_scores)
    rows = []
    for th in [0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]:
        recall = float((true_s >= th).mean()) if len(true_s) else 0.0
        keep = float((all_s >= th).mean()) if len(all_s) else 0.0
        rows.append({"阈值": th, "真同款通过率": round(recall, 3), "全量保留率": round(keep, 3)})
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--topn", type=int, default=5, help="每本网品召回候选数")
    ap.add_argument("--limit", type=int, default=0, help="召回候选上限（试跑）")
    ap.add_argument("--accept", type=float, default=0.7, help="精排采信阈值")
    args = ap.parse_args()

    excl = load_exclude()
    truth = load_truth()
    print(f"真条码验证集 {sum(len(v) for v in truth.values())} 对 / {len(truth)} 品", flush=True)

    from sentence_transformers import CrossEncoder
    model = CrossEncoder(MODEL)
    print("精排模型已加载", flush=True)

    for which, fn in [("jd", "对齐_向量_bl_jd.jsonl"), ("mc", "对齐_向量_bl_mc.jsonl")]:
        t0 = time.time()
        recall = load_recall(fn, args.topn, excl, args.limit)
        kept, stat = attr_filter(recall)
        print(f"\n[{which}] 召回品{len(recall)} → 属性过滤后 {len(kept)} 对 {dict(stat)}", flush=True)

        # 精排打分（分批）
        texts = [[o.get("a_name", ""), o.get("b_name", "")] for o, _, _ in kept]
        scores = []
        B = 256
        for i in range(0, len(texts), B):
            ss = model.predict(texts[i:i+B], batch_size=32)
            scores.extend([float(s) for s in ss])
            if (i // B) % 20 == 0:
                print(f"  精排 {i+len(ss)}/{len(texts)} {time.time()-t0:.0f}s", flush=True)

        # 真条码分数（本网品 + 其真猫超品，只有 mc 侧有效）
        truth_scores = []
        if which == "mc":
            tp = []
            for pb, items in truth.items():
                if pb in excl:
                    continue
                bl_name = next((o.get("a_name", "") for o, _, _ in kept if str(o["a"]) == pb), None)
                if not bl_name:
                    continue
                for pid, nm in items:
                    tp.append([bl_name, nm])
            if tp:
                ts = model.predict(tp, batch_size=32)
                truth_scores = [float(s) for s in ts]
                print(f"  真条码精排 {len(truth_scores)} 对, 均分 {sum(truth_scores)/len(truth_scores):.3f}", flush=True)

        # 写新对齐文件
        out = ALIGN_DIR / f"对齐_v2_bl_{which}.jsonl"
        tier_cnt = Counter()
        with out.open("w", encoding="utf-8") as f:
            for (o, aa, bb), s in zip(kept, scores):
                tier = "采信" if s >= args.accept else "边界"
                tier_cnt[tier] += 1
                f.write(json.dumps({
                    "a": o["a"], "b": o["b"], "sim": round(s, 4), "sim_vec": o["sim"],
                    "tier": tier, "rank": o["rank"], "conflict": "",
                    "a_name": o.get("a_name"), "b_name": o.get("b_name"),
                    "a_attr": aa, "b_attr": bb, "model": MODEL,
                }, ensure_ascii=False) + "\n")
        print(f"  输出 {out.name}: {len(kept)} 行 {dict(tier_cnt)}, 耗时 {time.time()-t0:.0f}s", flush=True)

        rows = calibrate(None, scores, truth, truth_scores)
        print("  阈值校准:", [(r["阈值"], r["真同款通过率"]) for r in rows], flush=True)
        if which == "mc":
            cp = HERE / "v2_阈值校准.json"
            json.dump(rows, open(cp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            print(f"  校准表: {cp.name}", flush=True)


if __name__ == "__main__":
    main()
