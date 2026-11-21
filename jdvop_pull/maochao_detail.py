#!/usr/bin/env python3
"""猫超产品详情批量拉取（含 SKU）— TOP alibaba.ascp.channel.distributor.product.detail。

list 接口只给 ID+标题，完整字段（品牌/类目/条码/零售价/税率/销售单位/属性/温控等）
需逐个回查详情接口（maochao-sdk 注释：铺货只带 ID，详情与价格需回查）。
输入：maochao_pull.py 产出的 raw jsonl（取最新一份）
输出：猫超产品详情_<n>个_<时间戳>.csv + 原始 jsonl
"""
import csv
import glob
import hashlib
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
import urllib.request

sys.path.insert(0, str(Path(__file__).resolve().parent))
from maochao_pull import load_creds, GATEWAY, CHANNEL, SUB_CHANNEL

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
METHOD = "alibaba.ascp.channel.distributor.product.detail"


def top_call(creds, biz):
    p = {"app_key": creds["app_key"], "method": METHOD, "session": creds["session"],
         "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "format": "json",
         "v": "2.0", "sign_method": "md5",
         "product_detail_request": json.dumps(biz, ensure_ascii=False)}
    src = creds["app_secret"] + "".join(f"{k}{p[k]}" for k in sorted(p)) + creds["app_secret"]
    p["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
    body = "&".join(f"{quote(k)}={quote(str(v), safe='')}" for k, v in p.items()).encode("utf-8")
    req = urllib.request.Request(GATEWAY, data=body,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def str_list(v):
    """TOP 列表包装 {string:[...]} → 列表。"""
    if isinstance(v, dict):
        return v.get("string") or []
    return v or []


def skus_of(d):
    psd = d.get("product_sku_detail") or {}
    if isinstance(psd, dict):
        return psd.get("product_sku_detail_d_t_o") or []
    return []


def flatten(d):
    pics = str_list(d.get("pictures"))
    barcodes = str_list(d.get("barcode_list"))
    props = str_list(d.get("properties"))
    mats = {}
    try:
        mats = json.loads(d.get("materials_map") or "{}")
    except Exception:
        pass
    skus = skus_of(d)
    return {
        "product_id": d.get("product_id"),
        "product_title": d.get("product_title"),
        "product_sub_title": d.get("product_sub_title"),
        "brand": d.get("brand"),
        "brand_id": d.get("brand_id"),
        "category": d.get("category"),
        "category_id": d.get("category_id"),
        "reserve_price": d.get("reserve_price"),
        "reserve_price_yuan": round(float(d.get("reserve_price") or 0) / 100, 2) if d.get("reserve_price") else None,
        "tax_rate": d.get("tax_rate"),
        "tax_code": d.get("tax_code"),
        "sale_unit": mats.get("sale_unit"),
        "barcode": barcodes[0] if barcodes else None,
        "temperature_requirement": d.get("temperature_requirement"),
        "sales_modes": ",".join(str_list(d.get("sales_modes"))),
        "supplier_user_id": d.get("supplier_user_id"),
        "picture": pics[0] if pics else None,
        "white_bg_picture": d.get("white_bg_picture"),
        "sku_count": len(skus),
        "properties": " | ".join(props),
    }


def flatten_sku(product_id, s):
    return {
        "product_id": product_id,
        "sku_id": s.get("sku_id"),
        "sku_properties": " | ".join(str_list(s.get("properties"))),
        "sku_reserve_price_yuan": round(float(s.get("reserve_price") or 0) / 100, 2) if s.get("reserve_price") else None,
        "sku_barcode": s.get("barcode"),
        "temperature_requirement": s.get("temperature_requirement"),
    }


def main():
    raw_files = sorted(glob.glob(str(OUT_DIR / "猫超产品_raw_*.jsonl")))
    if not raw_files:
        raise SystemExit("先跑 maochao_pull.py 拉产品列表")
    ids = []
    for l in open(raw_files[-1], encoding="utf-8"):
        if l.strip():
            ids.append(json.loads(l)["product_id"])
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else len(ids)
    ids = ids[:limit]
    print(f"回查详情 {len(ids)} 个产品", flush=True)

    creds = load_creds()
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    raw_out = OUT_DIR / f"猫超产品详情_raw_{stamp}.jsonl"
    csv_out = OUT_DIR / f"猫超产品详情_{len(ids)}个_{stamp}.csv"

    rows, sku_rows, fails = [], [], []
    with raw_out.open("w", encoding="utf-8") as raw:
        for i, pid in enumerate(ids, 1):
            biz = {"product_id": pid, "channel_code": CHANNEL,
                   "sub_channel_code": SUB_CHANNEL, "include_sku": True}
            for attempt in range(3):
                try:
                    resp = top_call(creds, biz)
                    break
                except Exception as exc:
                    if attempt == 2:
                        resp = None
                        last = exc
                    else:
                        time.sleep(1.5)
            if resp is None:
                fails.append({"product_id": pid, "error": str(last)})
                continue
            key = [k for k in resp if k.endswith("_response")]
            result = resp[key[0]].get("result", {}) if key else {}
            data = result.get("data")
            if not result.get("result_success") or not data:
                fails.append({"product_id": pid, "error": json.dumps(result, ensure_ascii=False)[:200]})
                continue
            raw.write(json.dumps(data, ensure_ascii=False) + "\n")
            rows.append(flatten(data))
            for s in skus_of(data):
                sku_rows.append(flatten_sku(pid, s))
            if i % 20 == 0 or i == len(ids):
                print(f"{i}/{len(ids)} 成功 {len(rows)} 失败 {len(fails)}", flush=True)
            time.sleep(0.3)

    fields = ["product_id", "product_title", "product_sub_title", "brand", "brand_id",
              "category", "category_id", "reserve_price", "reserve_price_yuan",
              "tax_rate", "tax_code", "sale_unit", "barcode",
              "temperature_requirement", "sales_modes",
              "supplier_user_id", "picture", "white_bg_picture", "sku_count", "properties"]
    with csv_out.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    sku_csv = OUT_DIR / f"猫超SKU明细_{len(sku_rows)}条_{stamp}.csv"
    with sku_csv.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["product_id", "sku_id", "sku_properties",
                                          "sku_reserve_price_yuan", "sku_barcode",
                                          "temperature_requirement"])
        w.writeheader()
        w.writerows(sku_rows)
    print(json.dumps({"详情成功": len(rows), "SKU条数": len(sku_rows), "失败": len(fails),
                      "csv": csv_out.name, "sku_csv": sku_csv.name, "raw": raw_out.name},
                     ensure_ascii=False, indent=2), flush=True)
    print(json.dumps({"详情成功": len(rows), "失败": len(fails), "csv": csv_out.name,
                      "raw": raw_out.name}, ensure_ascii=False, indent=2), flush=True)
    if fails:
        print("失败样本:", json.dumps(fails[:5], ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
