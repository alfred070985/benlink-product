#!/usr/bin/env python3
"""猫超（天猫超市企业购）渠道产品拉取 — TOP 开放协议直连。

依据 tvopx 仓库 maochao-sdk：
  - 接口 alibaba.ascp.channel.distributor.product.list（渠道产品批量查询）
  - 网关 https://eco.taobao.com/router/rest，MD5 签名 + session
  - 渠道码 tmallchaoshipm / tmcspm_qyg

凭证从 tvopx 的 application.yml 读取（可用环境变量 MAOCHAO_* 覆盖），不写死在代码。
只读拉取，分页小步、限速，不做任何写操作。
"""
import csv
import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
import urllib.request

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
YML = Path(os.environ.get("MAOCHAO_YML",
          Path.home() / "ai-workspace/projects/tvopx/tvopx-servicehost/src/main/resources/application.yml"))
GATEWAY = "https://eco.taobao.com/router/rest"
METHOD = "alibaba.ascp.channel.distributor.product.list"
CHANNEL = "tmallchaoshipm"
SUB_CHANNEL = "tmcspm_qyg"


def load_creds():
    text = YML.read_text(encoding="utf-8")
    def grab(key, env):
        v = os.environ.get(env)
        if v:
            return v
        m = re.search(rf"{key}:\s*\$\{{[^:]+\:(.+?)\}}", text)
        if not m:
            m = re.search(rf"{key}:\s*(\S+)", text)
        return m.group(1).strip() if m else None
    creds = {
        "app_key": grab("app-key", "MAOCHAO_APP_KEY"),
        "app_secret": grab("app-secret", "MAOCHAO_APP_SECRET"),
        "session": grab("session-key", "MAOCHAO_SESSION_KEY"),
    }
    missing = [k for k, v in creds.items() if not v]
    if missing:
        raise SystemExit(f"凭证缺失: {missing}（检查 {YML} 或环境变量）")
    return creds


def top_call(creds, biz_params):
    """TOP MD5 签名调用。biz_params 会作为整体 JSON 放到 method 对应的业务参数键下。"""
    sys_params = {
        "app_key": creds["app_key"],
        "method": METHOD,
        "session": creds["session"],
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "format": "json",
        "v": "2.0",
        "sign_method": "md5",
        "product_list_request": json.dumps(biz_params, ensure_ascii=False),
    }
    src = creds["app_secret"] + "".join(f"{k}{sys_params[k]}" for k in sorted(sys_params)) + creds["app_secret"]
    sys_params["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
    body = "&".join(f"{quote(k)}={quote(str(v), safe='')}" for k, v in sys_params.items()).encode("utf-8")
    req = urllib.request.Request(GATEWAY, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def flatten(p):
    """把产品对象摊平成 csv 行（字段名以 TOP 实际返回为准，snake_case）。"""
    pics = p.get("pictures") or {}
    pic = (pics.get("string") or [""])[0] if isinstance(pics, dict) else ""
    modes = p.get("sales_modes") or {}
    mode = ",".join(modes.get("string") or []) if isinstance(modes, dict) else ""
    return {
        "product_id": p.get("product_id"),
        "product_title": p.get("product_title"),
        "sales_modes": mode,
        "picture": pic,
        "gmt_modified": p.get("gmt_modified"),
    }


def extract_items(resp):
    """真实响应路径: alibaba_ascp_channel_distributor_product_list_response
       .product_list_response.data.products.products"""
    r = resp.get("alibaba_ascp_channel_distributor_product_list_response") or resp
    plr = r.get("product_list_response") or r.get("productListResponse") or {}
    data = plr.get("data") or {}
    prods = data.get("products") or {}
    items = prods.get("products") if isinstance(prods, dict) else prods
    if not items:
        items = data.get("product_list") or data.get("productList") or []
    total = data.get("total_count") or data.get("totalCount") or data.get("total")
    return items or [], total


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 100
    page_size = 20
    creds = load_creds()
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    raw_path = OUT_DIR / f"猫超产品_raw_{stamp}.jsonl"
    csv_path = OUT_DIR / f"猫超产品_{limit}个_{stamp}.csv"

    rows, page = [], 1
    with raw_path.open("w", encoding="utf-8") as raw:
        while len(rows) < limit:
            biz = {"channel_code": CHANNEL, "sub_channel_code": SUB_CHANNEL,
                   "page_no": page, "page_size": page_size, "is_scroll": False}
            resp = top_call(creds, biz)
            if "error_response" in resp:
                print("接口报错:", json.dumps(resp["error_response"], ensure_ascii=False)[:400], flush=True)
                break
            items, total = extract_items(resp)
            if not items:
                print(f"第{page}页无数据，共拉 {len(rows)} 个", flush=True)
                break
            for p in items:
                raw.write(json.dumps(p, ensure_ascii=False) + "\n")
                rows.append(flatten(p))
            print(f"page {page}: +{len(items)} 累计 {len(rows)} (total={total})", flush=True)
            if len(items) < page_size:
                break
            page += 1
            time.sleep(0.6)

    rows = rows[:limit]
    fields = ["product_id", "product_title", "sales_modes", "picture", "gmt_modified"]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    print(json.dumps({"拉取产品数": len(rows), "csv": csv_path.name, "raw": raw_path.name},
                     ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
