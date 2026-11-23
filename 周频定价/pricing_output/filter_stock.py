#!/usr/bin/env python3
"""对标准池候选商品批量查库存，筛出「四站都有库存(SaleQty>0)」的商品。

来源：archery 只读查 OpenAPI_Product.dbo.Product_Inventory（SQL Server 实例 256）
规则：IsDelete=0、Status=1、SaleQty>0，且 北京(1)/广州(2)/上海(3)/武汉(7) 四站都有记录
输出：标准池_四站全上架有库存_<日期>.jsonl
"""
import json
import subprocess
import time
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")
SRC = PRICING_DIR / f"标准池候选_四站全上架_{TODAY}.jsonl"
OUT = PRICING_DIR / f"标准池_四站全上架有库存_{TODAY}.jsonl"
LOG_OUT = PRICING_DIR / f"库存筛选说明_{TODAY}.json"

ARCHERY = "/Users/alfredmini/Benlink/ops-local/skills/benlai-trace-debug/scripts/archery/archery_client.py"
PY = "/Users/alfredmini/ai-workspace/local/venv/bin/python"
BATCH = 1000
SITES_SQL = "(1,2,3,7)"


def query_stock(ids):
    in_list = ",".join(str(i) for i in ids)
    sql = (
        "SELECT ProductSysNo FROM Product_Inventory "
        f"WHERE IsDelete=0 AND WebSiteSysNo IN {SITES_SQL} AND SaleQty>0 AND Status=1 "
        f"AND ProductSysNo IN ({in_list}) "
        "GROUP BY ProductSysNo HAVING COUNT(DISTINCT WebSiteSysNo)=4"
    )
    proc = subprocess.run(
        [PY, ARCHERY, "--action", "query", "--instance", "256",
         "--db", "OpenAPI_Product", "--sql", sql, "--limit", str(len(ids))],
        capture_output=True, text=True, timeout=120,
    )
    out = proc.stdout
    # 从 JSON 输出段解析；rows 为二维数组 [[ProductSysNo], ...]
    marker = "--- JSON 输出 ---"
    if marker not in out:
        raise RuntimeError(f"archery 输出异常: {out[-300:]}")
    payload = json.loads(out.split(marker, 1)[1])
    rows = payload.get("rows") or []
    return {r[0] for r in rows}


def main():
    t0 = time.time()
    candidates = [json.loads(l) for l in SRC.open(encoding="utf-8") if l.strip()]
    print(f"候选 {len(candidates)} 个", flush=True)

    stocked = {}
    fails = []
    for i in range(0, len(candidates), BATCH):
        batch = candidates[i:i + BATCH]
        ids = [c["productSysNo"] for c in batch]
        for attempt in range(3):
            try:
                ok = query_stock(ids)
                break
            except Exception as exc:
                if attempt == 2:
                    fails.append({"from": i, "error": str(exc)})
                    ok = set()
                else:
                    time.sleep(2)
        for c in batch:
            if c["productSysNo"] in ok:
                stocked[c["productSysNo"]] = c
        if (i // BATCH + 1) % 10 == 0 or i + BATCH >= len(candidates):
            print(f"batch {i//BATCH+1}/{(len(candidates)+BATCH-1)//BATCH}: 有库存累计 {len(stocked)}", flush=True)
        time.sleep(0.4)

    with OUT.open("w", encoding="utf-8") as f:
        for psys in sorted(stocked):
            f.write(json.dumps(stocked[psys], ensure_ascii=False) + "\n")

    log = {
        "筛选时间": datetime.now().isoformat(timespec="seconds"),
        "候选商品数": len(candidates),
        "四站都有库存商品数": len(stocked),
        "失败批": fails,
        "库存口径": "Product_Inventory: IsDelete=0, Status=1, SaleQty>0, 四站(北京1/广州2/上海3/武汉7)均有记录",
        "输出": str(OUT),
        "耗时秒": round(time.time() - t0, 1),
    }
    LOG_OUT.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(log, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
