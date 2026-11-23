#!/usr/bin/env python3
"""拉取分销池(ProductPoolType=1)在北京/广州/上海/武汉(1,2,3,7)四站的上架商品×站点关系，全量落 jsonl。

来源：vm-online.benlai.com /business-product-servicehost/ProductWebSiteApi/GetProductPager
筛选：COnlineStatus=1（=productWebSiteStatus=1 上架）
只读：不修改任何数据。路径由 __file__ 推导。
"""
import json
import time
from datetime import datetime
from pathlib import Path

import requests

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")
OUT = PRICING_DIR / f"源数据_分销池四站上架关系_{TODAY}.jsonl"
LOG_OUT = PRICING_DIR / f"取数说明_分销池四站上架关系_{TODAY}.json"

BASE = "https://vm-online.benlai.com"
URL = "/business-product-servicehost/ProductWebSiteApi/GetProductPager"
SITES = [1, 2, 3, 7]          # 北京、广州、上海、武汉
POOL_TYPE = 1                 # 分销商品池
PAGE_SIZE = 5000


def fetch_page(idx):
    payload = {
        "Query": {
            "SelectCount": idx == 1,
            "ProductName": "", "ProductId": "",
            "WebSiteSysNoList": SITES,
            "C1SysNo": -999999, "COnlineStatus": 1,
            "ExclusivePlatformProduct": False, "MerchantSysNo": -999999,
            "PageIndex": idx, "PageSize": PAGE_SIZE,
        },
        "ProductPoolType": POOL_TYPE, "FilterStatus": False,
    }
    for attempt in range(4):
        try:
            r = requests.post(BASE + URL, data={"request": json.dumps(payload)}, timeout=120)
            r.raise_for_status()
            body = r.json()
            if body.get("code") != 0:
                raise RuntimeError(body.get("msg"))
            return body["value"]
        except Exception as exc:
            if attempt == 3:
                raise
            time.sleep(2 * (attempt + 1))


def main():
    t0 = time.time()
    total_hint = None
    written = 0
    page = 1
    KEEP = ["sysNo", "productSysNo", "productBasicSysNo", "webSiteSysNo",
            "productWebSiteStatus", "basicPrice", "areaGroupSysNo", "areaGroupName",
            "isDelete", "lastUpdateTime"]
    with OUT.open("w", encoding="utf-8") as f:
        while True:
            v = fetch_page(page)
            if total_hint is None:
                total_hint = v.get("total") or 0
                print(f"total={total_hint}", flush=True)
            lst = v.get("list") or []
            if not lst:
                break
            for it in lst:
                f.write(json.dumps({k: it.get(k) for k in KEEP}, ensure_ascii=False) + "\n")
            written += len(lst)
            if page % 10 == 0 or len(lst) < PAGE_SIZE:
                print(f"page {page}: 累计 {written}", flush=True)
            if len(lst) < PAGE_SIZE:
                break
            page += 1
            time.sleep(0.15)

    log = {
        "pulled_at": datetime.now().isoformat(timespec="seconds"),
        "source": "vm-online ProductWebSiteApi/GetProductPager（只读分页查询）",
        "filter": {"ProductPoolType": 1, "WebSiteSysNoList": SITES, "COnlineStatus": 1},
        "sites": {"1": "北京", "2": "广州", "3": "上海", "7": "武汉"},
        "接口报告总数": total_hint,
        "实落条数": written,
        "read_only": True,
        "output": str(OUT),
        "elapsed_seconds": round(time.time() - t0, 1),
    }
    LOG_OUT.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print("完成:", json.dumps(log, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
