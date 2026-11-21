#!/usr/bin/env python3
"""按标准池 P 码批量回 ERP 取「商品×站点」成本明细（多站点聚合原材料）。

来源：本来生活 ERP 只读接口 ProLhotse/GetListByBasic（接口文档：仅对 product_channel_info / 价格表 SELECT）。
只读：不写库、不发事件、不调用任何更新接口。
输入：标准池_四站全上架有库存_<日期>.jsonl（productBasicSysNo 即 P 码数字部分）
输出：源数据_标准池ERP站点成本_<日期>.jsonl + 取数说明
"""
import json
import time
from datetime import datetime
from pathlib import Path

import requests

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(__import__("os").environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")
SRC = PRICING_DIR / f"标准池_四站全上架有库存_{TODAY}.jsonl"
OUTPUT = PRICING_DIR / f"源数据_标准池ERP站点成本_{TODAY}.jsonl"
META = PRICING_DIR / f"取数说明_标准池ERP站点成本_{TODAY}.json"

URL = "https://vm-online.benlai.com/erp-product-webapi/ProLhotse/GetListByBasic"
BATCH = 200
KEEP = [
    "ProductID", "ProductName", "ProductBasicSysNo", "ProductSysNo",
    "WebSiteSysNo", "ChannelSysNo", "SaleChannelSysNo", "Status", "IsDelete",
    "UnitCost", "DefaultPurchasePrice", "CurrentPrice", "BasicPrice",
    "IsDirectSend", "OriginDirectDelivery", "BenlaiCouldDirectDelivery",
]


def main():
    basics = []
    with SRC.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                basics.append(json.loads(line)["productBasicSysNo"])
    basics = sorted(set(basics))
    print(f"标准池 P 码 {len(basics)} 个", flush=True)

    session = requests.Session()
    t0 = time.time()
    total = 0
    failures = []
    nb = (len(basics) + BATCH - 1) // BATCH
    with OUTPUT.open("w", encoding="utf-8") as out:
        for i in range(0, len(basics), BATCH):
            batch = basics[i:i + BATCH]
            last_exc = None
            values = []
            for attempt in range(3):
                try:
                    resp = session.post(URL, data={"productBasicSysNoList": json.dumps(batch)}, timeout=180)
                    resp.raise_for_status()
                    payload = resp.json()
                    if payload.get("Code") != 0:
                        raise RuntimeError(payload.get("Msg"))
                    values = payload.get("Value") or []
                    break
                except Exception as exc:
                    last_exc = exc
                    time.sleep(1.5 * (attempt + 1))
            else:
                failures.append({"batch_from": i, "error": str(last_exc)})
                print(json.dumps({"batch": i // BATCH + 1, "error": str(last_exc)}, ensure_ascii=False), flush=True)
                continue
            for v in values:
                out.write(json.dumps({k: v.get(k) for k in KEEP}, ensure_ascii=False) + "\n")
            total += len(values)
            bno = i // BATCH + 1
            if bno % 25 == 0 or bno == nb:
                print(f"batch {bno}/{nb}: 累计 {total} 条, 用时 {time.time()-t0:.0f}s", flush=True)
            time.sleep(0.2)

    meta = {
        "pulled_at": datetime.now().isoformat(timespec="seconds"),
        "source": "本来生活 ERP 只读接口 ProLhotse/GetListByBasic（仅 SELECT product_channel_info / 价格表）",
        "read_only": True,
        "side_effects": "未写库、未发事件、未调用任何更新接口",
        "baseline": str(SRC),
        "baseline_pcodes": len(basics),
        "batch_size": BATCH,
        "records": total,
        "failed_batches": failures,
        "status_enum": {"-1": "作废", "0": "有效（未展示）", "1": "上架（展示）"},
        "fields": KEEP,
        "output": str(OUTPUT),
        "elapsed_seconds": round(time.time() - t0, 2),
    }
    META.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
