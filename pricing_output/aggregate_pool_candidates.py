#!/usr/bin/env python3
"""聚合四站上架关系 → 筛出「四站全部上架」的商品清单（标准池候选）。

输入：源数据_分销池四站上架关系_<日期>.jsonl（pull_pool_sites.py 产出）
规则：商品在 北京(1)/广州(2)/上海(3)/武汉(7) 四站都有 productWebSiteStatus=1 的记录
输出：标准池候选_<日期>.jsonl（每行一个商品，含四站 basicPrice）
"""
import json
from collections import defaultdict
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")
SRC = PRICING_DIR / f"源数据_分销池四站上架关系_{TODAY}.jsonl"
OUT = PRICING_DIR / f"标准池候选_四站全上架_{TODAY}.jsonl"
LOG_OUT = PRICING_DIR / f"聚合说明_标准池候选_{TODAY}.json"

SITES = {1: "北京", 2: "广州", 3: "上海", 7: "武汉"}


def main():
    rel = defaultdict(dict)   # productSysNo -> {webSiteSysNo: record}
    n = 0
    with SRC.open(encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            o = json.loads(line)
            n += 1
            if o.get("productWebSiteStatus") != 1 or o.get("isDelete"):
                continue
            rel[o["productSysNo"]][o["webSiteSysNo"]] = o

    qualified = []
    for psys, by_site in rel.items():
        if set(by_site.keys()) >= set(SITES.keys()):
            qualified.append({
                "productSysNo": psys,
                "productBasicSysNo": by_site[1]["productBasicSysNo"],
                "basicPrice": {str(s): by_site[s].get("basicPrice") for s in SITES},
                "areaGroupName": by_site[1].get("areaGroupName", ""),
            })

    qualified.sort(key=lambda x: x["productSysNo"])
    with OUT.open("w", encoding="utf-8") as f:
        for q in qualified:
            f.write(json.dumps(q, ensure_ascii=False) + "\n")

    log = {
        "聚合时间": datetime.now().isoformat(timespec="seconds"),
        "输入关系条数": n,
        "去重商品数": len(rel),
        "四站全上架商品数": len(qualified),
        "四站": SITES,
        "输出": str(OUT),
    }
    LOG_OUT.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(log, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
