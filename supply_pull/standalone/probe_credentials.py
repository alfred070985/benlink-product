#!/usr/bin/env python3
"""独立 VOP 账号：确认三串凭证各自的角色（纯只读探测，不写任何业务数据、不申请 token）。"""
import hashlib
import json
import time
from pathlib import Path

import requests

CONF = Path(__file__).with_name("credentials.json")
c = json.loads(CONF.read_text(encoding="utf-8"))
A, B, C = c["s1"], c["s2"], c["s3"]
URL = "https://api.jd.com/routerjson"


def call(method, biz, ak, sk, tk):
    params = {
        "method": method,
        "app_key": ak,
        "access_token": tk,
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
        "format": "json",
        "v": "2.0",
        "sign_method": "md5",
        "360buy_param_json": json.dumps(biz, ensure_ascii=False, separators=(",", ":")),
    }
    src = sk + "".join(k + str(params[k]) for k in sorted(params)) + sk
    params["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
    r = requests.post(URL, data=params, timeout=25)
    return r.status_code, r.text


def brief(text):
    t = text.replace("\n", " ")
    return t[:420]


combos = [
    ("s1=appKey s2=appSecret s3=token", A, B, C),
    ("s2=appKey s1=appSecret s3=token", B, A, C),
    ("s1=appKey s2=appSecret s1=token", A, B, A),
    ("s1=appKey s2=appSecret s2=token", A, B, B),
    ("s3=appKey s2=appSecret s1=token", C, B, A),
]

print("=== A. 用 getSkuPoolInfo 探测凭证组合 ===")
for label, ak, sk, tk in combos:
    try:
        status, text = call("jingdong.vop.goods.getSkuPoolInfo", {}, ak, sk, tk)
        print(f"{label} | HTTP {status} | {brief(text)}")
    except Exception as exc:
        print(f"{label} | 异常 {type(exc).__name__}: {str(exc)[:120]}")

print("\n=== B. 基准接口 getSkuDetailInfo（已知方法名可用） ===")
for label, ak, sk, tk in combos[:2]:
    try:
        status, text = call("jingdong.vop.goods.getSkuDetailInfo", {"skuId": 102215}, ak, sk, tk)
        print(f"{label} | HTTP {status} | {brief(text)}")
    except Exception as exc:
        print(f"{label} | 异常 {type(exc).__name__}: {str(exc)[:120]}")
