#!/usr/bin/env python3
"""探测：商品池全清单 + 其它编码类字段是否有官方中文字典接口。只读。"""
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
        "method": method, "app_key": app["app-key"], "access_token": token,
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
        "format": "json", "v": "2.0", "sign_method": "md5",
        "360buy_param_json": json.dumps(biz, ensure_ascii=False, separators=(",", ":")),
    }
    src = app["app-secret"] + "".join(k + str(params[k]) for k in sorted(params)) + app["app-secret"]
    params["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
    return requests.post(URL, data=params, timeout=20).text


def result(method, biz):
    try:
        txt = call(method, biz)
        j = json.loads(txt)
        if "error_response" in j:
            return "ERR " + j["error_response"].get("zh_desc", "")[:60]
        key = method.replace(".", "_") + "_responce"
        rpc = j.get(key, {}).get("openRpcResult", {})
        return f"code={rpc.get('resultCode')} {rpc.get('resultMessage')} data={json.dumps(rpc.get('result'), ensure_ascii=False)[:260]}"
    except Exception as exc:
        return "异常 " + str(exc)[:80]


print("=== 商品池全清单（示意：接口能列出账号下所有池） ===")
txt = call("jingdong.vop.goods.getSkuPoolInfo", {})
pools = json.loads(txt)["jingdong_vop_goods_getSkuPoolInfo_responce"]["openRpcResult"]["result"]["skuPoolList"]
print(f"共 {len(pools)} 个池")
for p in pools[:20]:
    print(" ", p.get("bizPoolId"), p.get("name"))

print("\n=== 其它编码/字典接口试探 ===")
trials = [
    ("jingdong.vop.goods.getTaxCodeList", {}),
    ("jingdong.vop.goods.getTaxCodeInfo", {"taxCode": "1080407020000000000"}),
    ("jingdong.vop.goods.getTaxInfo", {"taxCode": "1080407020000000000"}),
    ("jingdong.vop.goods.getBrandInfoList", {}),
    ("jingdong.vop.goods.getBrandList", {}),
    ("jingdong.vop.goods.getBrandInfo", {"brandId": 1}),
    ("jingdong.vop.goods.getReturnRuleInfoList", {}),
    ("jingdong.vop.goods.getAreaInfoList", {}),
    ("jingdong.vop.goods.getCategoryInfoList", {"categoryId": "6196,6198,11981,652,829,854,1319"}),
]
for method, biz in trials:
    print(f"{method} {json.dumps(biz, ensure_ascii=False)} | {result(method, biz)}")
