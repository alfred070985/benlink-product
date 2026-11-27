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
# 用户 2026-11-26 拍板：规格/件数写法不对称时，不能当"没信息"放过，须天然减分。
# 典型：本网"珍珠米5KG"(净含量base，无显式件数) vs 猫超"丝苗米5kgx4袋"(件数4、base空)，
# 精排 0.9432 误采信，折算价 24.18 远低于成本 40.67。扣 0.20 后降为边界。
SPEC_MISMATCH_PENALTY = 0.20
# 用户 2026-11-27 拍板：可比维度差异超过 18 倍 → 完全不当同品，直接排除；
# 18 倍以内（含）算同款不同包装，放行进入折算比价。典型：
#   网线 0.5米 vs 305米=610倍 → 排除；水 1瓶 vs 12瓶=12倍 → 可比；1 vs 24瓶=24倍 → 排除。
SPEC_RATIO_LIMIT = 18.0


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


def spec_asymmetry(aa, bb):
    """规格矛盾判定（用户 2026-11-26 拍板：规格对不上天然减分而非放过）。

    关键区分（2026-11-26 真条码 2539 对实测）：
      · 猫超品名普遍不写规格（如"蒙牛特仑苏营养纯牛奶苗条装"）——这是"信息缺失"，
        不是"规格矛盾"。若把这类一并减分，会误伤 61.8% 真同款，毁掉链路。
      · 真正矛盾 = 一方明确单件(有净含量base、无显式件数、count=1)，
        另一方明确多件(pack_explicit 且 count>1)。典型"珍珠米5KG vs 丝苗米5kgx4袋"。
    实测该精准规则对真同款误伤仅 3.5%，丝苗米case 命中。
    """
    a_single = bool(aa["base"]) and not aa["pack_explicit"] and aa["count"] == 1
    b_single = bool(bb["base"]) and not bb["pack_explicit"] and bb["count"] == 1
    a_multi = bool(aa["pack_explicit"]) and aa["count"] > 1
    b_multi = bool(bb["pack_explicit"]) and bb["count"] > 1
    return (a_single and b_multi) or (b_single and a_multi)


def ratio_over_limit(aa, bb):
    """18 倍规则（用户 2026-11-27 拍板）：两边都有同一可比维度且差异>18 倍 → 排除。

    只在两边同维度都有值时比，单边有值=信息缺失不算矛盾（避免误伤真同款）。
    维度来自 attr_extract.spec_values，已清洗营养含量/型号/单份克重/线芯直径。
    返回 (是否超限, 命中的维度与倍数描述)。
    """
    sa, sb = aa.get("specs") or {}, bb.get("specs") or {}
    for dim in ("count", "len", "vol"):
        va, vb = sa.get(dim) or [], sb.get(dim) or []
        if not va or not vb:
            continue
        for x in va:
            for y in vb:
                if x <= 0 or y <= 0:
                    continue
                ratio = max(x, y) / min(x, y)
                if ratio > SPEC_RATIO_LIMIT:
                    return True, f"{dim} {x:g}vs{y:g}={ratio:.0f}倍"
    return False, ""


def attr_filter(recall):
    """属性强约束（2026-11-27 版）：
      · 品牌冲突 → 直接否
      · 本网有明确件数、对方无件数无规格（佳沛 case）→ 直接否
      · 同维度差异>18 倍 → 直接否（用户拍板"超18倍不当同品"）
      · 件数/规格不等但≤18 倍 → 放行（旧版一律硬拦，误杀了 5L vs 2.5L 等同款）
      · 规格写法不对称（单边件数/单边规格，如对方面×4袋本网5KG）→ 减分 0.20
    返回保留对 (o, aa, bb, penalize) + 统计。
    """
    kept = []
    stat = Counter()
    for a, cs in recall.items():
        for o in cs:
            aa, bb = attrs_of(o.get("a_name")), attrs_of(o.get("b_name"))
            ba, bb_ = brand_tokens(o.get("a_name")), brand_tokens(o.get("b_name"))
            why = ""
            if ba and bb_ and not (ba & bb_):
                why = "品牌冲突"
            elif aa["pack_explicit"] and not bb["pack_explicit"] and not bb["base"]:
                # 用户 11-25 拍板延续：本网有明确件数、对方既无件数也无规格 → 不可比价，拦
                why = "对方无规格"
            else:
                over, detail = ratio_over_limit(aa, bb)
                if over:
                    why = "超18倍"
                    stat[detail] += 1
            if why:
                stat[why] += 1
            else:
                pen = spec_asymmetry(aa, bb)
                kept.append((o, aa, bb, pen))
                stat["规格不对称-减分" if pen else "保留"] += 1
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
        texts = [[o.get("a_name", ""), o.get("b_name", "")] for o, _, _, _ in kept]
        scores = []
        B = 256
        for i in range(0, len(texts), B):
            ss = model.predict(texts[i:i+B], batch_size=32)
            scores.extend([float(s) for s in ss])
            if (i // B) % 20 == 0:
                print(f"  精排 {i+len(ss)}/{len(texts)} {time.time()-t0:.0f}s", flush=True)

        # 规格不对称天然减分（用户 11-26 拍板）
        final_scores = [round(s - (SPEC_MISMATCH_PENALTY if pen else 0.0), 4)
                        for s, (_, _, _, pen) in zip(scores, kept)]

        # 真条码分数（本网品 + 其真猫超品，只有 mc 侧有效）
        truth_scores = []
        if which == "mc":
            tp = []
            for pb, items in truth.items():
                if pb in excl:
                    continue
                bl_name = next((o.get("a_name", "") for o, _, _, _ in kept if str(o["a"]) == pb), None)
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
            for (o, aa, bb, pen), s in zip(kept, final_scores):
                tier = "采信" if s >= args.accept else "边界"
                tier_cnt[tier] += 1
                f.write(json.dumps({
                    "a": o["a"], "b": o["b"], "sim": s, "sim_vec": o["sim"],
                    "tier": tier, "rank": o["rank"], "conflict": "",
                    "spec_penalty": SPEC_MISMATCH_PENALTY if pen else 0.0,
                    "a_name": o.get("a_name"), "b_name": o.get("b_name"),
                    "a_attr": aa, "b_attr": bb, "model": MODEL,
                }, ensure_ascii=False) + "\n")
        print(f"  输出 {out.name}: {len(kept)} 行 {dict(tier_cnt)}, 耗时 {time.time()-t0:.0f}s", flush=True)

        rows = calibrate(None, final_scores, truth, truth_scores)
        print("  阈值校准:", [(r["阈值"], r["真同款通过率"]) for r in rows], flush=True)
        if which == "mc":
            cp = HERE / "v2_阈值校准.json"
            json.dump(rows, open(cp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            print(f"  校准表: {cp.name}", flush=True)


if __name__ == "__main__":
    main()
