#!/usr/bin/env python3
"""探测京东 VOP 类目字典接口：确认能把 6196;6198;11981 这类类目 ID 翻成中文。

仅用之前从线上 Redis 缓存读出的只读 token 快照，不申请新 token、不写任何数据。
"""
import hashlib
import json
import time
from pathlib import Path

import requests
import yaml

CONF = Path("/tmp/benlai-online-conf/application.yml")
TOKENS = Path("/tmp/benlai-online-conf/tokens.json")
URL = "https://api.jd.com/routerjson"

cfg = yaml.safe_load(CONF.read_text(encoding="utf-8"))
app = cfg["jd-vop"]["sdk"]["apps"][0]
tokens = json.loads(TOKENS.read_text(encoding="utf-8"))
token = tokens["erp:third_vendor:jd:access_token:" + app["username"]]["token"]


def call(method, biz):
    params = {
        "method": method,
        "app_key": app["app-key"],
        "access_token": token,
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
        "format": "json",
        "v": "2.0",
        "sign_method": "md5",
        "360buy_param_json": json.dumps(biz, ensure_ascii=False, separators=(",", ":")),
    }
    src = app["app-secret"] + "".join(k + str(params[k]) for k in sorted(params)) + app["app-secret"]
    params["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
    r = requests.post(URL, data=params, timeout=20)
    return r.text[:600].replace("\n", " ")


methods = [
    "jingdong.vop.goods.getCategoryInfoList",
    "jingdong.vop.goods.getCategoryInfo",
    "jingdong.vop.goods.queryCategoryInfoList",
    "jingdong.vop.goods.getCategoryList",
    "jingdong.vop.goods.queryCategoryList",
    "jingdong.vop.goods.getCategoryById",
    "jingdong.vop.goods.getCategoryNameById",
    "jingdong.vop.goods.getSkuCategory",
    "jingdong.vop.goods.getAllCategoryInfo",
    "jingdong.vop.goods.getChildCategory",
    "jingdong.vop.goods.getSkuPoolInfo",
]

print("=== 方法名 + categoryId 试探 ===")
for m in methods:
    biz = {} if "SkuPoolInfo" in m else {"categoryId": 6196}
    try:
        print(f"{m} | {call(m, biz)}")
    except Exception as exc:
        print(f"{m} | 异常 {type(exc).__name__}: {str(exc)[:100]}")

print("\n=== 对 getCategoryInfoList 换参数试探 ===")
for biz in [
    {"categoryId": "6196,6198,11981"},
    {"categoryId": 6196},
    {"cid": 6196},
    {"parentId": 0},
    {"parentCid": 0},
    {"level": 1},
    {},
]:
    try:
        print(f"{json.dumps(biz, ensure_ascii=False)} | {call('jingdong.vop.goods.getCategoryInfoList', biz)}")
    except Exception as exc:
        print(f"{json.dumps(biz, ensure_ascii=False)} | 异常 {str(exc)[:100]}")
