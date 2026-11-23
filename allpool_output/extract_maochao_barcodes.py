#!/usr/bin/env python3
"""从猫超详情/SKU 产物提取条码索引：barcode → 猫超 product_id（SPU 级）。

输出：allpool_output/猫超条码索引.jsonl  每行 {"barcode": "692xxx", "product_id": "123", "sku_id": null}
      SKU 条码也收（sku_id 非空），对齐主用 SPU 条码，SKU 条码备用。
"""
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "猫超条码索引.jsonl"


def str_list(v):
    if isinstance(v, dict):
        return v.get("string") or []
    return v or []


def main():
    import glob
    from latest import latest
    n = 0
    seen = set()
    with OUT.open("w", encoding="utf-8") as f:
        # 周频：并集所有历史详情/SKU 文件（新品覆盖旧品），条码索引始终是全量的
        for pat, env in (("supply_pull/output/猫超详情_全量_*.jsonl", "ALLPOOL_MC_DETAIL"),
                         ("supply_pull/output/猫超SKU_全量_*.jsonl", "ALLPOOL_MC_SKU")):
            anchor = latest(pat, env)
            for fp in sorted(glob.glob(str(anchor.parent / pat.split("/")[-1]))):
                for l in open(fp, encoding="utf-8"):
                    if not l.strip():
                        continue
                    o = json.loads(l)
                    if "data" in o:
                        pid = o["product_id"]
                        for bc in str_list(o["data"].get("barcode_list")):
                            bc = str(bc).strip()
                            if bc and (bc, pid, None) not in seen:
                                seen.add((bc, pid, None))
                                f.write(json.dumps({"barcode": bc, "product_id": pid, "sku_id": None}, ensure_ascii=False) + "\n")
                                n += 1
                    else:
                        bc = str(o.get("barcode") or "").strip()
                        sid = str(o.get("sku_id") or "").strip()
                        if bc and sid and (bc, o["product_id"], sid) not in seen:
                            seen.add((bc, o["product_id"], sid))
                            f.write(json.dumps({"barcode": bc, "product_id": o["product_id"], "sku_id": sid}, ensure_ascii=False) + "\n")
                            n += 1
    print(f"条码索引 {n} 条 → {OUT.name}")


if __name__ == "__main__":
    main()
