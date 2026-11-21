#!/usr/bin/env python3
"""条码对齐：本网标准池 75,067 品 ↔ 猫超 SPU（条码精确匹配）。

逻辑：本网品的多条条码任一条命中猫超 SPU 条码（详情 barcode_list）即对齐。
输出：allpool_output/条码对齐_本网猫超.jsonl
      每行 {"pbasic":.., "maochao_pid":.., "barcode":..}
统计打印：对齐率、一对多分布。
"""
import json
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "条码对齐_本网猫超.jsonl"


def norm(bc):
    """条码规范化：去空白，去前导零的备用键由调用处决定。先精确串匹配。"""
    return str(bc).strip()


def main():
    # 猫超 SPU 条码索引（sku_id 为 null 的是 SPU 条码）
    mc_bc = {}
    for l in (HERE / "猫超条码索引.jsonl").open(encoding="utf-8"):
        if not l.strip():
            continue
        o = json.loads(l)
        if o.get("sku_id"):
            continue
        mc_bc.setdefault(norm(o["barcode"]), []).append(o["product_id"])

    n = hit = 0
    pairs = set()
    pair_bc = {}
    multi = 0
    no_bc = 0
    with OUT.open("w", encoding="utf-8") as f:
        for l in (HERE / "本网条码.jsonl").open(encoding="utf-8"):
            if not l.strip():
                continue
            o = json.loads(l)
            n += 1
            bcs = [norm(b) for b in (o.get("barcodes") or []) if str(b).strip()]
            if not bcs:
                no_bc += 1
                continue
            found = []
            for bc in bcs:
                for pid in mc_bc.get(bc, []):
                    found.append((pid, bc))
            if not found:
                continue
            hit += 1
            pids = {p for p, _ in found}
            if len(pids) > 1:
                multi += 1
            # 一对多：取第一个命中的；全部记录候选供上游复核
            for pid, bc in found:
                key = (o["pbasic"], pid)
                if key not in pairs:
                    pairs.add(key)
                    pair_bc[key] = bc
                    f.write(json.dumps({"pbasic": o["pbasic"], "maochao_pid": pid, "barcode": bc,
                                        "n_candidates": len(pids)}, ensure_ascii=False) + "\n")

    print(json.dumps({
        "本网品数": n, "无条码": no_bc,
        "对齐命中品数": hit, "对齐率": f"{hit/n:.1%}",
        "多候选品数(一对多)": multi,
        "对齐记录数(含多候选)": len(pairs),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
