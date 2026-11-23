#!/usr/bin/env python3
"""按标准池 productSysNo 批量查 product 表，补名称与类目编码（C1/C2/C3）。

来源：archery 只读查 benlai_business_product.product（MySQL 实例 147）
输出：源数据_标准池商品信息_<日期>.jsonl（每行：productSysNo, name, c1, c2, c3）
"""
import json
import subprocess
import time
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")
SRC = PRICING_DIR / f"标准池_四站全上架有库存_{TODAY}.jsonl"
OUT = PRICING_DIR / f"源数据_标准池商品信息_{TODAY}.jsonl"
LOG_OUT = PRICING_DIR / f"取数说明_标准池商品信息_{TODAY}.json"

ARCHERY = "/Users/alfredmini/Benlink/ops-local/skills/benlai-trace-debug/scripts/archery/archery_client.py"
PY = "/Users/alfredmini/ai-workspace/local/venv/bin/python"
BATCH = 800


def query_batch(ids):
    in_list = ",".join(str(i) for i in ids)
    sql = (f"SELECT SysNo, ProductName, C1SysNo, C2SysNo, C3SysNo, VendorName "
           f"FROM product WHERE SysNo IN ({in_list})")
    proc = subprocess.run(
        [PY, ARCHERY, "--action", "query", "--instance", "147",
         "--db", "benlai_business_product", "--sql", sql, "--limit", str(len(ids))],
        capture_output=True, text=True, timeout=120,
    )
    out = proc.stdout
    marker = "--- JSON 输出 ---"
    if marker not in out:
        raise RuntimeError(f"archery 输出异常: {out[-300:]}")
    payload = json.loads(out.split(marker, 1)[1])
    cols = payload.get("column_list") or []
    rows = payload.get("rows") or []
    return [dict(zip(cols, r)) for r in rows]


def main():
    t0 = time.time()
    ids = [json.loads(l)["productSysNo"] for l in SRC.open(encoding="utf-8") if l.strip()]
    print(f"待查 {len(ids)} 个", flush=True)
    total = 0
    fails = []
    nb = (len(ids) + BATCH - 1) // BATCH
    with OUT.open("w", encoding="utf-8") as f:
        for i in range(0, len(ids), BATCH):
            batch = ids[i:i + BATCH]
            for attempt in range(3):
                try:
                    rows = query_batch(batch)
                    break
                except Exception as exc:
                    if attempt == 2:
                        fails.append({"from": i, "error": str(exc)})
                        rows = []
                    else:
                        time.sleep(2)
            for r in rows:
                f.write(json.dumps({
                    "productSysNo": r.get("SysNo"),
                    "name": r.get("ProductName") or "",
                    "c1": r.get("C1SysNo"),
                    "c2": r.get("C2SysNo"),
                    "c3": r.get("C3SysNo"),
                    "vendor": r.get("VendorName") or "",
                }, ensure_ascii=False) + "\n")
            total += len(rows)
            bno = i // BATCH + 1
            if bno % 10 == 0 or bno == nb:
                print(f"batch {bno}/{nb}: 累计 {total}", flush=True)
            time.sleep(0.3)

    log = {
        "取数时间": datetime.now().isoformat(timespec="seconds"),
        "输入商品数": len(ids),
        "查到商品信息数": total,
        "失败批": fails,
        "输出": str(OUT),
        "耗时秒": round(time.time() - t0, 1),
    }
    LOG_OUT.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(log, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
