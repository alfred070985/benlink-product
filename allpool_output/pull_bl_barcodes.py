#!/usr/bin/env python3
"""archery 批量查本网标准池商品条码（product_barcode 表，P 码补零 10 位）。

输出：allpool_output/本网条码.jsonl  每行 {"pbasic": 585965, "pcode": "P0000585965", "barcodes": ["692xxx", ...]}
分批 IN 查询，断点续跑（重启跳过已查到的 pbasic）。
"""
import json
import subprocess
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
OUT = HERE / "本网条码.jsonl"
SRC = ROOT / "pricing_output" / "标准池_四站全上架有库存_2026-11-19.jsonl"
ARCHERY = "/Users/alfredmini/Benlink/ops-local/skills/benlai-trace-debug/scripts/archery/archery_client.py"
PY = "/Users/alfredmini/ai-workspace/local/venv/bin/python"
BATCH = 800


def query_batch(pcodes):
    in_list = ",".join(f"'{p}'" for p in pcodes)
    sql = (f"SELECT ProductID, BarCode FROM product_barcode "
           f"WHERE IsDelete=0 AND BarCode<>'' AND ProductID IN ({in_list})")
    proc = subprocess.run([PY, ARCHERY, "--action", "query", "--instance", "147",
                           "--db", "benlai_business_product", "--sql", sql,
                           "--limit", str(len(pcodes) * 3)],
                          capture_output=True, text=True, timeout=180)
    marker = "--- JSON 输出 ---"
    if marker not in proc.stdout:
        raise RuntimeError(f"archery 异常: {proc.stdout[-200:]}")
    payload = json.loads(proc.stdout.split(marker, 1)[1])
    rows = payload.get("rows") or []
    out = {}
    for r in rows:
        pid, bc = str(r[0]), str(r[1]).strip()
        if bc:
            out.setdefault(pid, []).append(bc)
    return out


def main():
    items = []
    for l in SRC.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            items.append((o["productBasicSysNo"], f"P{o['productBasicSysNo']:010d}"))
    done = set()
    if OUT.exists():
        for l in OUT.open(encoding="utf-8"):
            if l.strip():
                done.add(json.loads(l)["pbasic"])
    todo = [x for x in items if x[0] not in done]
    print(f"标准池 {len(items)}，已查 {len(done)}，待查 {len(todo)}", flush=True)

    t0 = time.time()
    nb = (len(todo) + BATCH - 1) // BATCH
    with OUT.open("a", encoding="utf-8") as f:
        for i in range(0, len(todo), BATCH):
            batch = todo[i:i + BATCH]
            for attempt in range(3):
                try:
                    res = query_batch([p for _, p in batch])
                    break
                except Exception:
                    if attempt == 2:
                        raise
                    time.sleep(3)
            got = set()
            for pb, pc in batch:
                bcs = res.get(pc) or []
                got.add(pb)
                f.write(json.dumps({"pbasic": pb, "pcode": pc, "barcodes": bcs}, ensure_ascii=False) + "\n")
            bno = i // BATCH + 1
            if bno % 10 == 0 or bno == nb:
                hit = sum(1 for pb, _ in batch if res.get(f"P{pb:010d}"))
                print(f"batch {bno}/{nb}: 累计 {len(done)+len(got)}，本批有条码 {hit}/{len(batch)}，{time.time()-t0:.0f}s", flush=True)
            done |= got
            time.sleep(0.3)
    have = sum(1 for _ in OUT.open(encoding='utf-8'))
    print(f"完成：{have} 个品的条码 → {OUT.name}", flush=True)


if __name__ == "__main__":
    main()
