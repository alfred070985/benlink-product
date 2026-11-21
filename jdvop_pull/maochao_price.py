#!/usr/bin/env python3
"""猫超全量价格回查（渠道供货价 + 券后价）— 3线程并发，断点续跑。

接口: alibaba.ascp.channel.distributor.price.get（price_type 单值，逐类型调用）
- CHANNEL_PRICE      渠道供货价（协议价，含税）——98% 覆盖
- COUPON_PRICE       券后价 ——98% 覆盖
- ADVISE_SALE_PRICE  建议零售价 —— 抽样 0% 覆盖（猫超未启用），不查，宽表留空列

产物: 猫超价格_全量_<日期>.jsonl  每行 {"product_id":..., "channel_price": 元, "coupon_price": 元,
      "channel_sku_n":..., "coupon_sku_n":...}（SPU 无值时取 SKU 最低价为该品代表价）
断点续跑: 重启扫描已有产物跳过已完成 product_id。
用法: python maochao_price.py
"""
import glob
import hashlib
import json
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
import urllib.request

sys.path.insert(0, str(Path(__file__).resolve().parent))
from maochao_pull import load_creds, GATEWAY, CHANNEL, SUB_CHANNEL

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
SLEEP = float(os.environ.get("MAOCHAO_SLEEP", "0.08"))
WORKERS = int(os.environ.get("MAOCHAO_WORKERS", "3"))
METHOD = "alibaba.ascp.channel.distributor.price.get"
PRICE_TYPES = ("CHANNEL_PRICE", "COUPON_PRICE")
_lock = threading.Lock()


def call(creds, biz):
    p = {"app_key": creds["app_key"], "method": METHOD, "session": creds["session"],
         "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "format": "json",
         "v": "2.0", "sign_method": "md5",
         "price_request": json.dumps(biz, ensure_ascii=False)}
    src = creds["app_secret"] + "".join(f"{k}{p[k]}" for k in sorted(p)) + creds["app_secret"]
    p["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
    body = "&".join(f"{quote(k)}={quote(str(v), safe='')}" for k, v in p.items()).encode("utf-8")
    req = urllib.request.Request(GATEWAY, data=body,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def parse_price(data):
    """返回 (SPU价元 or None, {sku_id: 价元})。"""
    def yuan(v):
        try:
            return round(float(v) / 100, 2) if v not in (None, "") else None
        except (TypeError, ValueError):
            return None
    spu = yuan((data.get("product_price") or {}).get("price"))
    skus = {}
    for sp in (data.get("sku_prices") or {}).get("top_channel_sku_price") or []:
        v = yuan((sp.get("sku_price") or {}).get("price"))
        if v is not None:
            skus[str(sp.get("sku_id"))] = v
    return spu, skus


def main():
    list_files = sorted(glob.glob(str(OUT_DIR / "猫超产品_list全量*.jsonl")))
    ids = [json.loads(l)["product_id"] for l in open(list_files[-1], encoding="utf-8") if l.strip()]
    print(f"基准产品 {len(ids)} 个", flush=True)

    today = datetime.now().strftime("%Y-%m-%d")
    price_path = OUT_DIR / f"猫超价格_全量_{today}.jsonl"
    fail_path = OUT_DIR / f"猫超价格_失败_{today}.jsonl"

    done = set()
    if price_path.exists():
        for l in price_path.open(encoding="utf-8"):
            if l.strip():
                done.add(json.loads(l)["product_id"])
    todo = [i for i in ids if i not in done]
    print(f"已完成 {len(done)}，待查 {len(todo)}", flush=True)

    creds = load_creds()
    t0 = time.time()
    cnt = {"ok": 0, "fail": 0}
    fp = price_path.open("a", encoding="utf-8")
    ff = fail_path.open("a", encoding="utf-8")

    def work(pid):
        out = {"product_id": pid}
        for pt in PRICE_TYPES:
            biz = {"product_id": pid, "channel_code": CHANNEL,
                   "sub_channel_code": SUB_CHANNEL, "sku_ids": [],
                   "price_type": pt, "sales_mode": "dealer"}
            spu = skus = None
            err = ""
            for attempt in range(6):
                try:
                    resp = call(creds, biz)
                    keys = [k for k in resp if k.endswith("_response")]
                    r = resp[keys[0]].get("result", {}) if keys else {}
                    if r.get("success") and r.get("data"):
                        spu, skus = parse_price(r["data"])
                        break
                    err = json.dumps(r, ensure_ascii=False)[:150]
                except Exception as exc:
                    err = str(exc)[:150]
                time.sleep(min(2 * (attempt + 1), 10))
            # SPU 无值时用 SKU 最低价为该品代表价
            rep = spu if spu is not None else (min(skus.values()) if skus else None)
            key = "channel_price" if pt == "CHANNEL_PRICE" else "coupon_price"
            out[key] = rep
            out[f"{key}_sku"] = skus or {}   # SKU 级价格明细（宽表按 sku_id 对齐）
            out.setdefault("_errs", [])
            if rep is None:
                out["_errs"].append(f"{pt}:{err[:60]}")
            time.sleep(SLEEP)
        with _lock:
            out["_errs"] = out["_errs"] or None
            if out["channel_price"] is None and out["coupon_price"] is None:
                ff.write(json.dumps({"product_id": pid, "error": out["_errs"]}, ensure_ascii=False) + "\n")
                cnt["fail"] += 1
            else:
                fp.write(json.dumps(out, ensure_ascii=False) + "\n")
                cnt["ok"] += 1
            i = cnt["ok"] + cnt["fail"]
            if i % 2000 == 0:
                fp.flush(); ff.flush()
                el = time.time() - t0
                rate = i / el
                eta = (len(todo) - i) / rate if rate else 0
                print(f"[{datetime.now():%H:%M:%S}] {i}/{len(todo)} 成功{cnt['ok']} 失败{cnt['fail']} "
                      f"{rate:.1f}/s ETA {eta/3600:.1f}h", flush=True)

    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        list(ex.map(work, todo))
    fp.close(); ff.close()
    print(json.dumps({"完成": cnt["ok"], "失败": cnt["fail"], "总数": len(ids)},
                     ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
