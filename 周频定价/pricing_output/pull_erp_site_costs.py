#!/usr/bin/env python3
"""一次性只读取数：按禄品 4940 个 P 码，拉「商品 × 站点」成本明细。

数据来源：本来生活 ERP 只读接口 ProLhotse/GetListByBasic。
按接口文档，该接口仅对 product_channel_info / 价格表执行 SELECT，不写任何数据。
本脚本只发只读请求、只落盘本地文件；不写库、不发事件、不调用任何更新接口。

路径全部由 __file__ 推导，可用环境变量覆盖：
  BENLINK_REPO  业务仓库根目录
  PRICING_DIR   本目录（输出目录）
"""
import json
import os
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import requests

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(os.environ.get("PRICING_DIR", HERE))
REPO = Path(os.environ.get("BENLINK_REPO", Path(__file__).resolve().parents[3]))
SOURCE_XLSX = REPO / "95012禄品业务" / "01_源数据" / "禄品4940_筛选表_2026-11-09.xlsx"
OUTPUT = PRICING_DIR / "源数据_ERP站点成本_2026-11-13.jsonl"
META = PRICING_DIR / "取数说明.json"

URL = "https://vm-online.benlai.com/erp-product-webapi/ProLhotse/GetListByBasic"
BATCH = 200
# 只保留参与定价/说明口径的字段，避免把商品详情长文本落盘
KEEP = [
    "ProductID", "ProductName", "ProductBasicSysNo", "ProductSysNo",
    "WebSiteSysNo", "ChannelSysNo", "SaleChannelSysNo", "Status", "IsDelete",
    "UnitCost", "DefaultPurchasePrice", "CurrentPrice", "BasicPrice",
    "IsDirectSend", "OriginDirectDelivery", "BenlaiCouldDirectDelivery",
]


def read_basics():
    wb = openpyxl.load_workbook(SOURCE_XLSX, read_only=True, data_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    header = list(rows[1])
    idx = header.index("商品编号")
    pcodes = [str(row[idx]).strip() for row in rows[2:] if row[idx]]
    wb.close()
    if len(pcodes) != len(set(pcodes)):
        raise RuntimeError("基准表 P 码有重复，停止取数")
    # 主商品编号 = P 码数字部分；仓库内两份成本数据已逐条验证过该映射
    basics = [int(p[1:]) for p in pcodes]
    return pcodes, basics


def main():
    pcodes, basics = read_basics()
    session = requests.Session()
    started = time.time()
    total = 0
    failures = []
    with OUTPUT.open("w", encoding="utf-8") as out:
        for i in range(0, len(basics), BATCH):
            batch = basics[i:i + BATCH]
            last_exc = None
            for attempt in range(3):
                try:
                    resp = session.post(
                        URL,
                        data={"productBasicSysNoList": json.dumps(batch)},
                        timeout=180,
                    )
                    resp.raise_for_status()
                    payload = resp.json()
                    if payload.get("Code") != 0:
                        raise RuntimeError(payload.get("Msg"))
                    values = payload.get("Value") or []
                    break
                except Exception as exc:  # 网络/网关瞬时失败，重试
                    last_exc = exc
                    time.sleep(1.5 * (attempt + 1))
            else:
                failures.append({"batch_from": i, "error": str(last_exc)})
                print(json.dumps({"batch": i // BATCH + 1, "error": str(last_exc)}, ensure_ascii=False), flush=True)
                continue
            for v in values:
                out.write(json.dumps({k: v.get(k) for k in KEEP}, ensure_ascii=False) + "\n")
            total += len(values)
            print(json.dumps({"batch": i // BATCH + 1, "requested": len(batch), "records": len(values)}, ensure_ascii=False), flush=True)
            time.sleep(0.2)

    meta = {
        "pulled_at": datetime.now().isoformat(timespec="seconds"),
        "source": "本来生活 ERP 只读接口 ProLhotse/GetListByBasic（接口文档：仅 SELECT product_channel_info / 价格表）",
        "read_only": True,
        "side_effects": "未写库、未发事件、未调用任何更新接口、未修改任何源表",
        "baseline": str(SOURCE_XLSX),
        "baseline_pcodes": len(pcodes),
        "batch_size": BATCH,
        "records": total,
        "failed_batches": failures,
        "status_enum": {"-1": "作废", "0": "有效（未展示）", "1": "上架（展示）"},
        "fields": KEEP,
        "output": str(OUTPUT),
        "elapsed_seconds": round(time.time() - started, 2),
    }
    META.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
