#!/usr/bin/env python3
"""给 v2 对齐结果回填品类字段（2026-11-27 用户拍板加品类维度）。

三源类目来源不同（详见 category_classify.py 头部注释）：
  · 本网 a=P码 → 本网商品底盘 → 三级类目ID → qc_db 字典中文名 → classify
  · 京东 b=sku → VOP 三段ID串查字典仅 4.5%（编码体系不同）→ 用品名 classify
  · 猫超 b=product_id → 猫超宽表中文 category → classify
回填字段：a_cat/a_cat_name/a_cat_src, b_cat/b_cat_name/b_cat_src,
          cat_conflict（是否命中硬互斥对）, combo

输入：对齐_v2_bl_jd.jsonl / 对齐_v2_bl_mc.jsonl
输出：同名覆盖（保留原字段，追加品类字段）；另存 品类回填统计.json
用法：python3 backfill_category.py [--dry-run N]
"""
import argparse
import csv
import json
from collections import Counter
from pathlib import Path

import openpyxl

from category_classify import classify, cat_conflict, is_combo, CATEGORY_KEYWORDS
from latest import latest

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent

CAT_NAMES = {cid: name for cid, (name, _) in CATEGORY_KEYWORDS.items()}
CAT_NAMES[0] = "其他"
DICT = HERE / "类目字典_qc_db_2026-11-27.json"


def load_cat_dict():
    d = json.loads(DICT.read_text(encoding="utf-8"))
    return {int(k): v["name"] for k, v in d["by_id"].items()}


def load_bl_cat():
    """本网 P码 → (一级,二级,三级) 类目ID。"""
    p = latest("本网商品底盘_*.xlsx")
    if not p:
        raise FileNotFoundError("本网商品底盘 not found")
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    hdr = next(rows)
    idx = {h: i for i, h in enumerate(hdr)}
    out = {}
    for r in rows:
        try:
            out[str(r[idx["商品编号(P码)"]])] = (r[idx["一级类目"]], r[idx["二级类目"]], r[idx["三级类目"]])
        except (IndexError, TypeError):
            continue
    return out


def load_mc_cat():
    """猫超 product_id → 中文类目名。"""
    p = latest("猫超商品宽表SKU级_*.csv")
    if not p:
        raise FileNotFoundError("猫超商品宽表 not found")
    out = {}
    for r in csv.DictReader(open(p, encoding="utf-8-sig")):
        pid = str(r.get("product_id", "")).strip()
        if pid and pid not in out:
            out[pid] = (r.get("category") or "").strip()
    return out


def resolve_cat(ids, cat_dict, name):
    """类目ID组 → 中文名（取最末级非空的）→ classify。"""
    cn = ""
    for v in reversed(ids):
        try:
            cn = cat_dict.get(int(v), "")
        except (ValueError, TypeError):
            cn = str(v or "").strip()
        if cn:
            break
    text = cn or name
    r = classify(text)
    return r, cn, ("字典" if cn else "品名")


def enrich(o, which, bl_cat, mc_cat, cat_dict):
    """给单条对齐记录补品类字段（in-place）。精排流水线与回填工具共用，保证重跑不丢品类。"""
    a, b = str(o["a"]), str(o["b"])
    a_name, b_name = o.get("a_name", ""), o.get("b_name", "")
    # 本网：P码 → 三级类目ID → 字典
    ids = bl_cat.get(a)
    if ids:
        ar, acn, asrc = resolve_cat(ids, cat_dict, a_name)
    else:
        ar, acn, asrc = classify(a_name), "", "品名"
    # 对照侧
    if which == "mc":
        bcn = mc_cat.get(b, "")
        br = classify(bcn) if bcn else classify(b_name)
        bsrc = "字典" if bcn else "品名"
    else:  # jd：类目ID不通，用品名
        bcn, br, bsrc = "", classify(b_name), "品名"
    combo = is_combo(a_name) or is_combo(b_name)
    conflict = cat_conflict(a_name, b_name, combo)
    o.update({
        "a_cat": ar["cat"], "a_cat_name": CAT_NAMES[ar["cat"]],
        "a_cat_word": ar["matched"], "a_cat_cn": acn, "a_cat_src": asrc,
        "b_cat": br["cat"], "b_cat_name": CAT_NAMES[br["cat"]],
        "b_cat_word": br["matched"], "b_cat_cn": bcn, "b_cat_src": bsrc,
        "cat_conflict": conflict, "combo": combo,
    })
    return o, asrc, bsrc, conflict, combo


def process(which, bl_cat, mc_cat, cat_dict, dry=0):
    fn = HERE / f"对齐_v2_bl_{which}.jsonl"
    rows = [json.loads(l) for l in fn.open(encoding="utf-8") if l.strip()]
    if dry:
        rows = rows[:dry]
    stat = Counter()
    out_rows = []
    for o in rows:
        _, asrc, bsrc, conflict, combo = enrich(o, which, bl_cat, mc_cat, cat_dict)
        stat[f"a_src_{asrc}"] += 1
        stat[f"b_src_{bsrc}"] += 1
        stat["品类冲突" if conflict else "品类不冲突"] += 1
        if combo:
            stat["组合礼盒"] += 1
        out_rows.append(o)

    if not dry:
        with fn.open("w", encoding="utf-8") as f:
            for o in out_rows:
                f.write(json.dumps(o, ensure_ascii=False) + "\n")
    return out_rows, stat


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", type=int, default=0)
    args = ap.parse_args()
    cat_dict = load_cat_dict()
    bl_cat = load_bl_cat()
    mc_cat = load_mc_cat()
    print(f"字典 {len(cat_dict)} 类目, 本网底盘 {len(bl_cat)} P码, 猫超 {len(mc_cat)} 品", flush=True)
    total = Counter()
    for which in ("jd", "mc"):
        rows, stat = process(which, bl_cat, mc_cat, cat_dict, args.dry_run)
        print(f"[{which}] {len(rows)} 行 {dict(stat)}", flush=True)
        total.update(stat)
    if not args.dry_run:
        json.dump(dict(total), open(HERE / "品类回填统计.json", "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
        print("落盘 品类回填统计.json", flush=True)


if __name__ == "__main__":
    main()
