#!/usr/bin/env python3
"""加白名单后的一键验证：确认三串凭证各自角色 + 类目字典接口 + 账号下商品池清单。

纯只读探测：不写任何业务数据、不修改凭证文件、不申请新通行证、不碰本来生活线上环境。
结论只打印，是否采纳由用户决定。
"""
import argparse
import hashlib
import json
import os
import time
from pathlib import Path

import requests

CONF = Path(__file__).with_name("credentials.json")
URL = "https://api.jd.com/routerjson"
WHITELIST_IPS = ("39.102.184.125", "39.102.184.126", "39.102.184.127",
                 "39.102.184.96", "8.153.82.69")

ap = argparse.ArgumentParser(description="加白名单/切换出网通道后的一键验证")
ap.add_argument("--proxy", default="", help="出网通道，如 socks5://127.0.0.1:1080 或 http://host:port")
args = ap.parse_args()
for key in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY",
            "http_proxy", "https_proxy", "all_proxy"):
    if args.proxy:
        os.environ[key] = args.proxy
    else:
        os.environ.pop(key, None)
if args.proxy:
    print(f"已启用出网通道：{args.proxy.split('@')[-1]}")
try:
    exit_ip = requests.get("https://api.ipify.org", timeout=10).text.strip()
except Exception as exc:
    exit_ip = f"未知（{str(exc)[:50]}）"
print(f"参考出口 IP（本机代理视角）：{exit_ip}")
print("注意：本机有隧道分流，京东实际看到的 IP 以下面第 1 步返回的 code 80 信息为准。\n")

c = json.loads(CONF.read_text(encoding="utf-8"))
A, B, C = c["s1"], c["s2"], c["s3"]


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
    return r.text


def parse(text):
    try:
        j = json.loads(text)
    except Exception:
        return ("非JSON", text[:200])
    if "error_response" in j:
        e = j["error_response"]
        return ("ERR", f"code={e.get('code')} {e.get('zh_desc','')[:120]}")
    key = next((k for k in j if k.endswith("_responce")), None)
    if not key:
        return ("未知结构", text[:200])
    rpc = j[key].get("openRpcResult", {})
    if str(rpc.get("resultCode")) not in ("0000", "0"):
        return ("业务失败", f"code={rpc.get('resultCode')} {rpc.get('resultMessage')}")
    return ("OK", rpc.get("result"))


combos = [
    ("s1=应用编号 / s2=应用密钥 / s3=通行证", A, B, C),
    ("s1=应用编号 / s2=应用密钥 / s2=通行证(试错)", A, B, B),
    ("s2=应用编号 / s1=应用密钥 / s3=通行证(试错)", B, A, C),
]

print("=== 1. 凭证组合验证（getSkuPoolInfo） ===")
winner = None
for label, ak, sk, tk in combos:
    try:
        status, data = parse(call("jingdong.vop.goods.getSkuPoolInfo", {}, ak, sk, tk))
        print(f"  {label} -> {status} | {str(data)[:260]}")
        if status == "OK" and isinstance(data, dict):
            winner = (label, ak, sk, tk, data)
    except Exception as exc:
        print(f"  {label} -> 异常 {type(exc).__name__}: {str(exc)[:120]}")

if not winner:
    msg = str(parse(call("jingdong.vop.goods.getSkuPoolInfo", {}, A, B, C))[1])
    if "白名单" in msg or "IP:" in msg:
        print("\n结论：仍然被 IP 白名单拦截，白名单尚未生效（或放行的不是当前出口 IP）。")
    elif "token" in msg.lower() or "通行" in msg or "无效" in msg:
        print("\n结论：白名单已放行，但 s3 不是有效通行证 —— 这套账号需要走「换发通行证」路径，")
        print("      还缺 VOP 登录用户名 / 登录密码 / 回调地址，请补齐后我再继续。")
    elif "签名" in msg:
        print("\n结论：白名单已放行，但签名不过 —— 第二串可能不是应用密钥，请核对。")
    else:
        print("\n结论：尚未通过，明细见上。请按错误信息核对凭证。")
    raise SystemExit(0)

label, ak, sk, tk, data = winner
pools = data.get("skuPoolList") or []
print(f"\n=== 2. 账号下商品池：共 {len(pools)} 个 ===")
for p in pools:
    print(f"  {p.get('bizPoolId')}  {p.get('name')}")

print("\n=== 3. 类目字典接口验证 ===")
status, cat = parse(call("jingdong.vop.goods.getCategoryInfoList",
                        {"categoryId": "6196,6198,11981"}, ak, sk, tk))
print(f"  {status} | {str(cat)[:300]}")

print("\n=== 4. 结论 ===")
print(f"  可用组合：{label}")
print("  s1 = 应用编号(appKey) / s2 = 应用密钥(appSecret) / s3 = 通行证(accessToken)")
print("  建议把 credentials.json 的 credential_mode 设为 'token'、user_decided 设为 true 后再跑主工具。")
print("  （本脚本不修改任何文件，是否采纳由你确认）")
