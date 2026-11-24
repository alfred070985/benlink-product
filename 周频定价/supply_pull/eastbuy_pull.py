#!/usr/bin/env python3
"""东方甄选 B2B Open API 商品拉取（只读）。

鉴权：X-Signature = Base64(HMAC-SHA256(appSecret, appKey+timestamp+nonce+path+rawBody))
GET 无 body 时 rawBody 固定为 "{}"。path 不含域名和 query。

用法：
  python eastbuy_pull.py [page_size]     全量分页拉取 /products（默认 pageSize=100）

凭证从环境变量读（不硬编码）：
  EASTBUY_APP_KEY / EASTBUY_APP_SECRET
"""
import base64
import hashlib
import hmac
import json
import os
import random
import string
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

BASE = "https://care.eastbuy.com"
PATH = "/api/open/v1/products"
HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
TODAY = datetime.now().strftime("%Y-%m-%d")


def creds():
    k, s = os.environ.get("EASTBUY_APP_KEY"), os.environ.get("EASTBUY_APP_SECRET")
    if not k or not s:
        raise SystemExit("缺少 EASTBUY_APP_KEY / EASTBUY_APP_SECRET 环境变量")
    return k, s


def sign(app_key, app_secret, path, raw_body):
    ts = str(int(time.time()))
    nonce = "".join(random.choices(string.ascii_letters + string.digits, k=16))
    msg = f"{app_key}{ts}{nonce}{path}{raw_body}"
    sig = base64.b64encode(hmac.new(app_secret.encode(), msg.encode(), hashlib.sha256).digest()).decode()
    return ts, nonce, sig


def get(path, params):
    app_key, app_secret = creds()
    ts, nonce, sig = sign(app_key, app_secret, path, "{}")
    qs = "&".join(f"{k}={urllib.parse.quote(str(v))}" for k, v in params.items())
    req = urllib.request.Request(f"{BASE}{path}?{qs}", method="GET", headers={
        "X-App-Key": app_key, "X-Timestamp": ts, "X-Nonce": nonce, "X-Signature": sig,
    })
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def main():
    page_size = int(sys.argv[1]) if len(sys.argv) > 1 else 100
    raw_path = OUT_DIR / f"东方甄选商品_{TODAY}.jsonl"
    page, n, total = 1, 0, None
    t0 = time.time()
    with raw_path.open("w", encoding="utf-8") as f:
        while True:
            resp = get(PATH, {"page": page, "pageSize": page_size})
            if resp.get("code") != 0:
                print(f"第{page}页失败: {resp.get('bizCode')} {resp.get('message')} requestId={resp.get('requestId')}")
                break
            data = resp.get("data") or {}
            items = data.get("list") or []
            total = data.get("total", total)
            for it in items:
                f.write(json.dumps(it, ensure_ascii=False) + "\n")
            n += len(items)
            if page % 20 == 0 or len(items) < page_size:
                print(f"page {page}: 累计 {n}/{total}  {time.time()-t0:.0f}s", flush=True)
            if not items or len(items) < page_size:
                break
            page += 1
            time.sleep(0.15)
    print(json.dumps({"拉取": n, "total_hint": total, "文件": raw_path.name}, ensure_ascii=False))


if __name__ == "__main__":
    import urllib.parse
    main()
