#!/usr/bin/env python3
"""货盘一选品重筛：在 75,067 池上加两条剔除规则（用户 2026-11-23 拍板）。

规则：
- 仓发 + 冷链(DeliveryTemperature 101/102) + 成本×1.12 < 30 元 → 剔除
- 仓发 + 常温(100) + 成本×1.12 < 10 元 → 剔除
- 直发品不受影响；无成本品保留（定价时本就不出价）

步骤：
1. archery 批量拉 75,067 品的 DeliveryTemperature（product 表）
2. 按拍板成本口径（build_standard_pool_pricing.resolve_cost）算成本
3. 应用剔除规则，产出新池 + 剔除清单
输出：
  标准池_选品v2_2026-11-23.jsonl     新池（含 temperature/is_cold/剔除标记字段）
  剔除清单_选品v2_2026-11-23.jsonl   被剔除的品+原因
"""
import json
import os
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "pricing_output"))
from build_standard_pool_pricing import resolve_cost, load_costs

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
from latest import latest
SRC = latest("pricing_output/标准池_四站全上架有库存_*.jsonl", "ALLPOOL_POOL")
COST = latest("pricing_output/源数据_标准池ERP站点成本_*.jsonl", "ALLPOOL_COST")
OUT_POOL = HERE / f"标准池_选品v2_{datetime.now():%Y-%m-%d}.jsonl"
OUT_CUT = HERE / f"剔除清单_选品v2_{datetime.now():%Y-%m-%d}.jsonl"
ARCHERY = os.environ.get("BENLAI_ARCHERY_CLIENT",
    str(Path.home() / "Benlink/ops-local/skills/benlai-trace-debug/scripts/archery/archery_client.py"))
PY = os.environ.get("WEEKLY_PY", str(Path.home() / "ai-workspace/local/venv/bin/python"))
BATCH = 800

WAREHOUSE_MID = 1.12   # 仓发中档加价（拍板）
COLD_LIMIT = 30.0      # 仓发冷链：加价后低于此价剔除
NORMAL_LIMIT = 10.0    # 仓发常温：加价后低于此价剔除


def pull_temperatures(psys_list):
    temps = {}
    for i in range(0, len(psys_list), BATCH):
        batch = psys_list[i:i + BATCH]
        in_list = ",".join(str(x) for x in batch)
        sql = f"SELECT SysNo, DeliveryTemperature FROM product WHERE SysNo IN ({in_list})"
        for attempt in range(3):
            try:
                proc = subprocess.run([PY, ARCHERY, "--action", "query", "--instance", "147",
                                       "--db", "benlai_business_product", "--sql", sql,
                                       "--limit", str(len(batch))],
                                      capture_output=True, text=True, timeout=180)
                payload = json.loads(proc.stdout.split("--- JSON 输出 ---", 1)[1])
                for r in payload.get("rows") or []:
                    temps[int(r[0])] = r[1]
                break
            except Exception:
                if attempt == 2:
                    raise
                time.sleep(3)
        if (i // BATCH + 1) % 20 == 0:
            print(f"温控 {len(temps)}/{len(psys_list)}", flush=True)
        time.sleep(0.25)
    return temps


def main():
    pool = [json.loads(l) for l in SRC.open(encoding="utf-8") if l.strip()]
    print(f"标准池 {len(pool)}", flush=True)
    temps = pull_temperatures([o["productSysNo"] for o in pool])
    print(f"温控拉到 {len(temps)}", flush=True)

    # 成本（拍板口径）
    cost_map = {}
    recs = {}
    for l in COST.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            recs[o["ProductBasicSysNo"]].append(o) if o["ProductBasicSysNo"] in recs else recs.setdefault(o["ProductBasicSysNo"], []).append(o)
    for pb, rs in recs.items():
        c, basis, direct = resolve_cost(rs)
        cost_map[pb] = (c, direct)

    kept, cut = [], []
    stat = Counter()
    for o in pool:
        pb = o["productBasicSysNo"]
        temp = temps.get(o["productSysNo"])
        is_cold = temp in (101, 102)
        cost, direct = cost_map.get(pb, (None, None))
        price_mid = round(cost * (1 + WAREHOUSE_MID), 2) if cost else None
        reason = None
        if direct is False:  # 仓发
            if price_mid is not None:
                if is_cold and price_mid < COLD_LIMIT:
                    reason = f"仓发冷链·中档价{price_mid}<{COLD_LIMIT}"
                elif not is_cold and price_mid < NORMAL_LIMIT:
                    reason = f"仓发常温·中档价{price_mid}<{NORMAL_LIMIT}"
        o2 = dict(o)
        o2["temperature"] = temp
        o2["is_cold"] = is_cold
        o2["ship"] = "直发" if direct else ("仓发" if direct is False else "未知")
        o2["cost"] = cost
        o2["mid_price"] = price_mid
        if reason:
            o2["剔除原因"] = reason
            cut.append(o2)
            stat[reason.split("·")[0]] += 1
        else:
            kept.append(o2)
    with OUT_POOL.open("w", encoding="utf-8") as f:
        for o in kept:
            f.write(json.dumps(o, ensure_ascii=False) + "\n")
    with OUT_CUT.open("w", encoding="utf-8") as f:
        for o in cut:
            f.write(json.dumps(o, ensure_ascii=False) + "\n")
    print(json.dumps({"原池": len(pool), "剔除": len(cut), "新池": len(kept),
                      "剔除构成": dict(stat)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
