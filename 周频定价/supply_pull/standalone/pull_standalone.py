#!/usr/bin/env python3
"""独立 JD VOP 账号：全量商品拉取 + 中文增强（本地只读工具）

用户已拍板的口径：
  - 范围：该账号下全部商品池，用京东接口自动发现，不依赖本来生活线上配置。
  - 中文：一律以新增列追加，原有 32 列字段值一个都不改。
  - 只有能从京东官方类目字典查到的 ID 才翻成中文；查不到就留空并标记来源。
  - 数字状态码不自行翻译，只输出取值清单交用户决定。

硬约束（写死，不通过参数绕过）：
  1. 凭证角色必须由用户本人确认：credentials.json 的 credential_mode + user_decided
     未确认时脚本直接拒绝运行，绝不推断第三串是通行证还是密码。
  2. 只用本目录 credentials.json，不读、不连、不借用本来生活线上 Apollo / Redis / 数据库。
  3. 不写任何生产库、不发商品变更事件、不调用线上 sync/all / sync/sku 接口。
  4. token 模式下不调用任何授权/换发接口；token 失效即停止并提示，不自动申请。
  5. 凭证不写死在代码里、不打印、不进日志。
"""
import argparse
import base64
import csv
import hashlib
import json
import os
import re
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

import requests

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"
CRED = ROOT / "credentials.json"
JD_URL = "https://api.jd.com/routerjson"
OAUTH_CODE_URL = "https://open-oauth.jd.com/oauth2/authorizeForVOP"
OAUTH_TOKEN_URL = "https://open-oauth.jd.com/oauth2/access_token"
PAGE_SIZE = 100
CATEGORY_BATCH = 120
PAGE_SLEEP = 0.12

# 与上一份（本来生活）产出完全一致的 32 列，顺序不变，方便对照
BASE_FIELDS = [
    "account", "pool_id", "sku_id", "sku_name", "brand_name", "category", "sale_unit",
    "sale_state", "sku_state", "is_can_vat", "ban_cause", "return_rule_type",
    "sale_price", "jd_price", "market_price", "naked_price", "tax_rate_percentage",
    "tax_code", "seo_model", "lowest_buy", "product_area", "color", "size", "weight",
    "length", "width", "height", "self_sell_type", "logistics_type",
    "image_primary", "image_count", "captured_at",
]
# 新增列：ID/编码 -> 中文，以及接口里本来就有、上次没落表的中文字段
EXTRA_FIELDS = [
    "pool_name",
    "spu_id", "spu_name",
    "c1_id", "c1_name", "c2_id", "c2_name", "c3_id", "c3_name", "category_full_name",
    "category_dict_status",
    "return_rule_str", "tax_price", "has_promotion",
    "sku_type", "giveaway_type", "save_energy_type",
    "no_reason_to_return", "output_vat", "pre_days",
    "warrant_desc", "category_attr", "param_group_attr",
]
CSV_FIELDS = BASE_FIELDS + EXTRA_FIELDS

STATUS_CODE_FIELDS = [
    "sale_state", "sku_state", "is_can_vat", "return_rule_type", "self_sell_type",
    "logistics_type", "sku_type", "giveaway_type", "save_energy_type",
    "no_reason_to_return", "output_vat", "has_promotion",
]


WHITELIST_IPS = ("39.102.184.125", "39.102.184.126", "39.102.184.127",
                 "39.102.184.96", "8.153.82.69")


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def apply_proxy(proxy):
    """让全部 HTTPS 请求走指定通道（HTTP 或 SOCKS5 均可）。

    不指定时清除环境里注入的本地代理，让流量走系统真实路由 —— 这样连上公司 VPN
    之后出口才会变成 VPN 的出口，而不是被本机代理固定到别的地址。
    """
    keys = ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY",
            "http_proxy", "https_proxy", "all_proxy")
    if not proxy:
        for key in keys:
            os.environ.pop(key, None)
        log("未指定出网通道：清除本机代理，直连走系统路由（连 VPN 后即为 VPN 出口）")
        return
    for key in keys:
        os.environ[key] = proxy
    log(f"已启用出网通道：{proxy.split('@')[-1]}")


def jd_seen_ip():
    """直接问京东：它实际看到的来源 IP 是多少（白名单校验以它为准）。

    本机有隧道分流时，api.ipify.org 看到的 IP 与京东看到的可能不同，不能混用。
    """
    try:
        cfg = json.loads(CRED.read_text(encoding="utf-8"))
        params = {
            "method": "jingdong.vop.goods.getSkuPoolInfo", "app_key": cfg["s1"],
            "access_token": "probe",
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
            "format": "json", "v": "2.0", "sign_method": "md5",
            "360buy_param_json": "{}",
        }
        src = cfg["s2"] + "".join(k + str(params[k]) for k in sorted(params)) + cfg["s2"]
        params["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
        text = requests.post(JD_URL, data=params, timeout=15).text
        m = re.search(r"IP:([\d.]+)", text)
        if m:
            return m.group(1)
        if "error_response" not in text:
            return "已放行"
    except Exception as exc:
        return f"未知（{str(exc)[:50]}）"
    return ""


def check_exit_ip():
    """打印京东实际会看到的来源 IP，并对照白名单。"""
    try:
        ip = requests.get("https://api.ipify.org", timeout=10).text.strip()
    except Exception as exc:
        ip = f"未知（{str(exc)[:60]}）"
    jd_ip = jd_seen_ip()
    if jd_ip:
        log(f"本机代理出口 IP：{ip}　|　京东实际看到的 IP：{jd_ip}")
        hit = "已在该账号的 IP 白名单中" if jd_ip in WHITELIST_IPS else (
            "京东已放行" if jd_ip == "已放行" else "不在白名单中，京东会拦截 code 80")
        log(f"白名单判定：{hit}")
        return jd_ip
    hit = "已在该账号的 IP 白名单中" if ip in WHITELIST_IPS else "不在白名单中，京东会拦截"
    log(f"当前出口 IP：{ip} —— {hit}")
    return ip


def s(v):
    return "" if v is None else str(v)


def load_credentials():
    """GUARDRAIL：凭证角色必须由用户确认，禁止自动推断。"""
    cfg = json.loads(CRED.read_text(encoding="utf-8"))
    mode = cfg.get("credential_mode", "undecided")
    if mode == "undecided" or not cfg.get("user_decided", False):
        raise RuntimeError(
            "GUARDRAIL 拦截：第三串凭证的角色尚未由你确认。\n"
            "  请先在 credentials.json 明确写 credential_mode（token / oauth）并把 user_decided 设为 true。\n"
            "  脚本不会替你判断它是通行证、登录密码还是用户名。\n"
            "  可先运行 verify_access.py 做只读探测辅助判断（探测结果不自动写入）。"
        )
    if mode == "token":
        token = cfg.get("access_token") or ""
        if not token:
            raise RuntimeError("GUARDRAIL 拦截：token 模式但未显式填写 access_token，不从 s1/s2/s3 猜测。")
        return cfg, token
    if mode == "oauth":
        for k in ("username", "password", "redirect_uri", "rsa_private_key", "app_key", "app_secret"):
            if not cfg.get(k):
                raise RuntimeError(f"GUARDRAIL 拦截：oauth 模式缺少显式配置的 {k}")
        return cfg, obtain_token_oauth(cfg)
    raise RuntimeError(f"GUARDRAIL 拦截：未知的 credential_mode={mode!r}")


def rsa_private_encrypt(msg: bytes, private_key_b64: str) -> bytes:
    """复刻项目 RsaCoderUtil.encryptByPrivateKey：PKCS#1 v1.5 type1 padding + 私钥幂运算。"""
    from cryptography.hazmat.primitives.serialization import load_der_private_key
    key = load_der_private_key(base64.b64decode(private_key_b64), password=None)
    k = key.key_size // 8
    em = b"\x00\x01" + b"\xff" * (k - 3 - len(msg)) + b"\x00" + msg
    nums = key.private_numbers()
    c = pow(int.from_bytes(em, "big"), nums.d, nums.n)
    return c.to_bytes(k, "big")


def obtain_token_oauth(cfg):
    """仅在 oauth 模式且用户填齐凭证时调用；不缓存、不写入凭证文件。"""
    md5pwd = hashlib.md5(cfg["password"].encode("utf-8")).hexdigest()
    enc = base64.b64encode(rsa_private_encrypt(md5pwd.encode("ascii"), cfg["rsa_private_key"])).decode("ascii")
    url = (f"{OAUTH_CODE_URL}?app_key={cfg['app_key']}&response_type=code"
           f"&redirect_uri={quote(cfg['redirect_uri'], safe='')}&username={quote(cfg['username'], safe='')}"
           f"&password={quote(enc, safe='')}&scope=snsapi_base")
    r = requests.get(url, allow_redirects=False, timeout=30)
    if r.status_code != 302:
        raise RuntimeError(f"换发通行证第一步失败，HTTP {r.status_code}")
    code = parse_qs(urlparse(r.headers.get("location", "")).query).get("code", [""])[0]
    if not code:
        raise RuntimeError("换发通行证第一步未返回 code")
    r2 = requests.get(
        f"{OAUTH_TOKEN_URL}?app_key={cfg['app_key']}&app_secret={cfg['app_secret']}"
        f"&code={quote(code, safe='')}&grant_type=authorization_code", timeout=30)
    j = r2.json()
    token = j.get("access_token") or j.get("accessToken") or ""
    if not token:
        raise RuntimeError(f"换发通行证第二步未返回通行证：{str(j)[:160]}")
    log("oauth 模式：已换发通行证（不写入凭证文件）")
    return token


class JdVopClient:
    def __init__(self, cfg, token):
        self.app_key = cfg["app_key"]
        self.app_secret = cfg["app_secret"]
        self.token = token

    def call(self, method, biz, retries=3):
        last = None
        for attempt in range(1, retries + 1):
            try:
                params = {
                    "method": method, "app_key": self.app_key, "access_token": self.token,
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
                    "format": "json", "v": "2.0", "sign_method": "md5",
                    "360buy_param_json": json.dumps(biz, ensure_ascii=False, separators=(",", ":")),
                }
                src = self.app_secret + "".join(k + str(params[k]) for k in sorted(params)) + self.app_secret
                params["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
                resp = requests.post(JD_URL, data=params, timeout=30)
                resp.raise_for_status()
                payload = resp.json()
                if "error_response" in payload:
                    e = payload["error_response"]
                    raise RuntimeError(f"code={e.get('code')} {e.get('zh_desc','')[:80]}")
                rpc = payload[method.replace(".", "_") + "_responce"]["openRpcResult"]
                if str(rpc.get("resultCode")) != "0":
                    raise RuntimeError(f"code={rpc.get('resultCode')} {rpc.get('resultMessage')}")
                return rpc.get("result")
            except Exception as exc:
                last = exc
                if attempt < retries:
                    time.sleep(0.4 * attempt)
        raise RuntimeError(str(last)[:200])


def discover_pools(client):
    result = client.call("jingdong.vop.goods.getSkuPoolInfo", {})
    pools = [{"id": s(p.get("bizPoolId")), "name": s(p.get("name"))}
             for p in (result or {}).get("skuPoolList", []) if p.get("bizPoolId")]
    if not pools:
        raise RuntimeError("该账号未返回任何商品池")
    return pools


def save_pools(pools):
    with (OUT / "商品池清单.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["pool_id", "pool_name"])
        w.writeheader()
        for p in pools:
            w.writerow({"pool_id": p["id"], "pool_name": p["name"]})


def build_category_dict(client, ids):
    """京东官方类目字典批量查 ID -> 中文；查不到留空，不猜。"""
    dict_path = OUT / "类目字典.json"
    cat = json.loads(dict_path.read_text(encoding="utf-8")) if dict_path.exists() else {}
    todo = [i for i in sorted(ids) if i and i not in cat]
    log(f"类目字典：需要查 {len(todo)} 个 ID，已有 {len(cat)} 条")
    for n in range(0, len(todo), CATEGORY_BATCH):
        batch = todo[n:n + CATEGORY_BATCH]
        try:
            rows = client.call("jingdong.vop.goods.getCategoryInfoList",
                               {"categoryId": ",".join(batch)}) or []
            for row in rows:
                cat[s(row.get("categoryId"))] = {
                    "name": s(row.get("categoryName")),
                    "level": s(row.get("categoryLevel")),
                    "parent_id": s(row.get("parentId")),
                }
        except Exception as exc:
            log(f"类目字典批次失败（留空标记）：{str(exc)[:120]}")
        time.sleep(0.1)
    dict_path.write_text(json.dumps(cat, ensure_ascii=False, indent=1), encoding="utf-8")
    with (OUT / "类目字典.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["category_id", "category_name", "category_level", "parent_id"])
        w.writeheader()
        for cid, info in sorted(cat.items()):
            w.writerow({"category_id": cid, "category_name": info["name"],
                       "category_level": info["level"], "parent_id": info["parent_id"]})
    return cat


def translate_category(raw, cat):
    """6196;6198;11981 -> 一级/二级/三级 ID + 中文名；缺字典就留空并标记来源。"""
    parts = [p.strip() for p in s(raw).split(";") if p.strip()]
    out = {k: "" for k in ("c1_id", "c1_name", "c2_id", "c2_name", "c3_id", "c3_name",
                           "category_full_name", "category_dict_status")}
    if not parts:
        out["category_dict_status"] = "empty"
        return out
    names, missing = [], []
    for idx, cid in enumerate(parts[:3], start=1):
        out[f"c{idx}_id"] = cid
        info = cat.get(cid)
        if info and info.get("name"):
            out[f"c{idx}_name"] = info["name"]
            names.append(info["name"])
        else:
            missing.append(cid)
    out["category_full_name"] = " > ".join(names) if not missing and len(names) == len(parts[:3]) else ""
    if missing:
        out["category_dict_status"] = "missing:" + ",".join(missing)
    elif len(parts) > 3:
        out["category_dict_status"] = "more_than_3_levels"
    else:
        out["category_dict_status"] = "ok"
    return out


def compact_attrs(value, limit=600):
    """把类目属性/参数组拼成可读中文，长文本截断，完整内容留在 raw.jsonl。"""
    if not value:
        return ""
    try:
        if isinstance(value, list):
            chunks = []
            for item in value:
                if isinstance(item, dict):
                    name = item.get("cateAttrName") or item.get("attrName") or item.get("groupName") or ""
                    vals = item.get("cateAttrValList") or item.get("attrValueList") or item.get("paramList") or ""
                    if isinstance(vals, list):
                        vals = ",".join(s(x) for x in vals)
                    chunks.append(f"{name}:{vals}" if name else s(vals))
                else:
                    chunks.append(s(item))
            text = " | ".join(chunks)
        else:
            text = s(value)
        return text[:limit]
    except Exception:
        return s(value)[:limit]


def flatten(rec, cat):
    """保持上一份 32 列取值口径不变，再追加中文列。"""
    detail = rec.get("detail") or {}
    price = rec.get("price") or {}
    image = rec.get("image") or {}
    sale = rec.get("sale") or {}
    state = rec.get("state") or {}
    images = image.get("skuImageList") or []
    primary = ""
    for item in images:
        if item.get("isPrimary") == 1 or item.get("position") == 0:
            primary = item.get("shortPath") or item.get("imagePath") or ""
            break
    if not primary and images:
        primary = images[0].get("shortPath") or ""

    row = {
        "account": rec.get("account", ""),
        "pool_id": rec.get("pool_id", ""),
        "sku_id": rec.get("sku_id", ""),
        "sku_name": detail.get("skuName") or sale.get("skuName") or "",
        "brand_name": detail.get("brandName") or "",
        "category": detail.get("category") or "",
        "sale_unit": detail.get("saleUnit") or "",
        "sale_state": sale.get("saleState", ""),
        "sku_state": state.get("skuState", detail.get("skuState", "")),
        "is_can_vat": sale.get("isCanVat", ""),
        "ban_cause": sale.get("banCause", ""),
        "return_rule_type": sale.get("returnRuleType", ""),
        "sale_price": price.get("salePrice", ""),
        "jd_price": price.get("jdPrice", ""),
        "market_price": price.get("marketPrice", ""),
        "naked_price": price.get("nakedPrice", ""),
        "tax_rate_percentage": price.get("taxRatePercentage", detail.get("taxRatePercentage", "")),
        "tax_code": detail.get("taxCode", ""),
        "seo_model": detail.get("seoModel", ""),
        "lowest_buy": detail.get("lowestBuy", ""),
        "product_area": detail.get("productArea", ""),
        "color": detail.get("color", ""),
        "size": detail.get("size", ""),
        "weight": detail.get("weight", ""),
        "length": detail.get("length", ""),
        "width": detail.get("width", ""),
        "height": detail.get("height", ""),
        "self_sell_type": detail.get("selfSellType", sale.get("selfType", "")),
        "logistics_type": detail.get("logisticsType", ""),
        "image_primary": primary,
        "image_count": len(images),
        "captured_at": rec.get("captured_at", ""),
        # ---- 新增列 ----
        "pool_name": rec.get("pool_name", ""),
        "spu_id": detail.get("spuId", ""),
        "spu_name": detail.get("spuName", ""),
        "return_rule_str": sale.get("returnRuleStr", ""),
        "tax_price": price.get("taxPrice", ""),
        "has_promotion": price.get("hasPromotion", ""),
        "sku_type": detail.get("skuType", ""),
        "giveaway_type": detail.get("giveawayType", ""),
        "save_energy_type": detail.get("saveEnergyType", ""),
        "no_reason_to_return": sale.get("noReasonToReturn", ""),
        "output_vat": sale.get("outputVat", ""),
        "pre_days": sale.get("preDays", ""),
        "warrant_desc": detail.get("warrantDesc", ""),
        "category_attr": compact_attrs(detail.get("categoryAttrList")),
        "param_group_attr": compact_attrs(detail.get("paramGroupAttrList")),
    }
    row.update(translate_category(row["category"], cat))
    return row


def load_progress():
    p = OUT / "progress.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def save_progress(data):
    p = OUT / "progress.json"
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(p)


def record_failure(fail_f, pool, sku, stage, err):
    fail_f.write(json.dumps({"pool_id": pool, "sku_id": sku, "stage": stage,
                            "error": str(err)[:300], "ts": datetime.now().isoformat()},
                           ensure_ascii=False) + "\n")
    fail_f.flush()


def pull(client, pools, account, sample, max_pages, workers):
    raw_path = OUT / "raw.jsonl"
    fail_path = OUT / "failures.jsonl"
    ids_path = OUT / "category_ids.json"
    progress = load_progress()
    ids = set(json.loads(ids_path.read_text(encoding="utf-8"))) if ids_path.exists() else set()
    total_new = 0
    failures = 0

    with raw_path.open("a", encoding="utf-8") as raw_f, fail_path.open("a", encoding="utf-8") as fail_f:
        for pool in (pools[:1] if sample else pools):
            pid, pname = pool["id"], pool["name"]
            st = progress.get(pid, {})
            if st.get("done"):
                log(f"池 {pid}（{pname}）已完成，跳过")
                continue
            offset, pages = st.get("offset", 0), st.get("pages", 0)
            log(f"池 {pid}（{pname}）从 offset={offset} 开始")
            while True:
                if max_pages and pages >= max_pages:
                    break
                try:
                    page = client.call("jingdong.vop.goods.querySkuByPage",
                                       {"bizPoolId": pid, "offset": offset, "pageSize": PAGE_SIZE})
                except Exception as exc:
                    record_failure(fail_f, pid, offset, "querySkuByPage", exc)
                    raise
                skus = (page or {}).get("skus") or []
                sku_csv = ",".join(str(x) for x in skus)
                sale_map, price_map, image_map, state_map = {}, {}, {}, {}
                if skus:
                    for method, target in [
                        ("jingdong.vop.goods.checkSkuSaleList", sale_map),
                        ("jingdong.vop.goods.getSkuImageList", image_map),
                        ("jingdong.vop.goods.getSkuStateList", state_map),
                    ]:
                        try:
                            rows = client.call(method, {"skuId": sku_csv}) or []
                            for r in rows:
                                if r and r.get("skuId") is not None:
                                    target[int(r["skuId"])] = r
                        except Exception as exc:
                            record_failure(fail_f, pid, "", method, exc)
                    can_sale = [x for x in skus if sale_map.get(int(x), {}).get("saleState") == 1]
                    if can_sale:
                        try:
                            rows = client.call("jingdong.vop.goods.getSellPrice",
                                               {"skuId": ",".join(str(x) for x in can_sale)}) or []
                            for r in rows:
                                if r and r.get("skuId") is not None:
                                    price_map[int(r["skuId"])] = r
                        except Exception as exc:
                            record_failure(fail_f, pid, "", "getSellPrice", exc)

                    def fetch_detail(sku):
                        try:
                            return sku, client.call("jingdong.vop.goods.getSkuDetailInfo",
                                                    {"skuId": int(sku)}), None
                        except Exception as exc:
                            return sku, None, exc

                    details = {}
                    with ThreadPoolExecutor(max_workers=workers) as ex:
                        for fut in as_completed([ex.submit(fetch_detail, x) for x in skus]):
                            sku, detail, err = fut.result()
                            if err:
                                record_failure(fail_f, pid, sku, "getSkuDetailInfo", err)
                            else:
                                details[int(sku)] = detail

                    captured = datetime.now().isoformat(timespec="seconds")
                    for sku in skus:
                        sid = int(sku)
                        detail = details.get(sid) or {}
                        for cid in s(detail.get("category")).split(";"):
                            if cid.strip():
                                ids.add(cid.strip())
                        rec = {
                            "account": account, "pool_id": pid, "pool_name": pname, "sku_id": sid,
                            "captured_at": captured,
                            "detail": detail, "price": price_map.get(sid),
                            "image": image_map.get(sid), "sale": sale_map.get(sid),
                            "state": state_map.get(sid),
                        }
                        raw_f.write(json.dumps(rec, ensure_ascii=False) + "\n")
                        total_new += 1
                    raw_f.flush()

                pages += 1
                next_offset = (page or {}).get("offset", offset + len(skus))
                remain = (page or {}).get("remainPage", 0)
                progress[pid] = {"offset": next_offset, "pages": pages, "done": remain == 0,
                                "skus": progress.get(pid, {}).get("skus", 0) + len(skus)}
                save_progress(progress)
                ids_path.write_text(json.dumps(sorted(ids), ensure_ascii=False), encoding="utf-8")
                log(f"  池 {pid} 第 {pages} 页：{len(skus)} 个 SKU，本次累计 {total_new}，剩余页 {remain}")
                if sample or remain == 0:
                    break
                offset = next_offset
                time.sleep(PAGE_SLEEP)
            if sample:
                break
    if fail_path.exists():
        with fail_path.open(encoding="utf-8") as f:
            failures = sum(1 for _ in f)
    return total_new, failures


def finalize(cat):
    """从 raw.jsonl 逐行生成最终 CSV（含中文列）+ 对账 + 状态码取值清单。"""
    raw_path = OUT / "raw.jsonl"
    csv_path = OUT / "goods.csv"
    status_counts = {k: Counter() for k in STATUS_CODE_FIELDS}
    seen = set()
    dup = 0
    rows = 0
    with raw_path.open(encoding="utf-8") as raw_f, \
            csv_path.open("w", encoding="utf-8-sig", newline="") as csv_f:
        writer = csv.DictWriter(csv_f, fieldnames=CSV_FIELDS, extrasaction="ignore")
        writer.writeheader()
        for line in raw_f:
            if not line.strip():
                continue
            rec = json.loads(line)
            row = flatten(rec, cat)
            writer.writerow(row)
            rows += 1
            key = (row["pool_id"], row["sku_id"])
            if key in seen:
                dup += 1
            else:
                seen.add(key)
            for k in STATUS_CODE_FIELDS:
                status_counts[k][s(row.get(k, ""))] += 1
    with (OUT / "状态码取值清单.txt").open("w", encoding="utf-8") as f:
        f.write("说明：以下只列出接口返回的数字状态码取值与出现次数，不翻译、不裁决。\n")
        f.write("如果要把某个状态码翻成中文，请确认口径后我再作为新增列追加，原值保持不变。\n\n")
        for k, counter in status_counts.items():
            f.write(f"[{k}] 共 {len(counter)} 种取值\n")
            for val, n in counter.most_common():
                f.write(f"  {val or '(空)'} : {n}\n")
            f.write("\n")
    return {"csv_rows": rows, "duplicate_pool_sku": dup,
            "category_dict_rows": len(cat)}


def main():
    ap = argparse.ArgumentParser(description="独立 VOP 账号全量拉取 + 中文增强")
    ap.add_argument("--sample", action="store_true", help="小样本：第一个池只拉 1 页")
    ap.add_argument("--max-pages", type=int, default=0, help="每个池最多页数（0=不限）")
    ap.add_argument("--workers", type=int, default=8, help="商品详情并发数")
    ap.add_argument("--finalize-only", action="store_true", help="不拉取，只重建 CSV/字典/报告")
    ap.add_argument("--proxy", default="", help="出网通道，如 socks5://127.0.0.1:1080 或 http://host:port")
    ap.add_argument("--check-ip-only", action="store_true", help="只查当前出口 IP 是否在白名单，不需要凭证")
    args = ap.parse_args()

    OUT.mkdir(parents=True, exist_ok=True)
    apply_proxy(args.proxy)
    exit_ip = check_exit_ip()
    if args.check_ip_only:
        return
    if exit_ip not in WHITELIST_IPS:
        log("提醒：当前出口不在白名单，京东会返回 code 80；请先切换出网通道。")

    cfg, token = load_credentials()
    client = JdVopClient(cfg, token)
    account = cfg.get("account_label", "独立VOP账号")

    if not args.finalize_only:
        pools = discover_pools(client)
        save_pools(pools)
        log(f"账号下共 {len(pools)} 个商品池，开始拉取")
        new_rows, failures = pull(client, pools, account, args.sample,
                                 1 if args.sample else args.max_pages, args.workers)
        log(f"本次新增 {new_rows} 条原始记录，失败记录 {failures} 条")
    else:
        log("跳过拉取，只重建最终表与报告")

    ids_path = OUT / "category_ids.json"
    ids = set(json.loads(ids_path.read_text(encoding="utf-8"))) if ids_path.exists() else set()
    cat = build_category_dict(client, ids)
    report = finalize(cat)
    report.update({
        "finished_at": datetime.now().isoformat(timespec="seconds"),
        "account": account,
        "credential_mode": cfg.get("credential_mode"),
        "scope": "该账号下全部商品池（getSkuPoolInfo 自动发现）",
        "chinese_rule": "仅京东官方类目字典可翻译；其它 ID/状态码保留原值并标记来源",
        "side_effects": "未写生产库、未发商品变更事件、未调用线上同步接口、未借用线上缓存 token",
    })
    (OUT / "对账.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    with (OUT / "goods.csv").open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        n = 0
        cat_ok = 0
        cat_missing = 0
        for row in reader:
            n += 1
            if row["category_dict_status"] == "ok":
                cat_ok += 1
            elif row["category_dict_status"].startswith("missing"):
                cat_missing += 1
    lines = [
        "独立 JD VOP 账号 全量商品数据（中文增强版）",
        "=" * 50,
        f"账号：{account}",
        f"数据行数：{n}",
        f"列数：{len(CSV_FIELDS)}（原 32 列 + 新增 {len(EXTRA_FIELDS)} 列）",
        f"商品池：见 商品池清单.csv",
        f"类目字典：{len(cat)} 条，见 类目字典.csv",
        f"类目三级中文完整：{cat_ok} 行；字典缺失而留空并标记：{cat_missing} 行",
        f"同池 SKU 重复：{report['duplicate_pool_sku']} 条",
        "",
        "中文列说明：",
        "  c1_name/c2_name/c3_name/category_full_name 来自京东官方类目字典接口，",
        "  只有查到才填；查不到留空并在 category_dict_status 标记 missing:<ID>。",
        "  pool_name 来自京东商品池接口；spu_name / return_rule_str / warrant_desc /",
        "  category_attr / param_group_attr 本来就是接口返回的中文字段，此次一并补进表。",
        "  数字状态码（销售状态、是否开票等）一律保留原值，不自行翻译；",
        "  各状态实际取值见 状态码取值清单.txt，需要中文口径请确认后再加列。",
        "  税收分类编码 taxCode 无京东官方字典接口，保留原编码。",
        "",
        "原 32 列取值口径与上一份（本来生活）产出完全一致，未做任何改写。",
        "完整原始记录（含图文详情等长文本）见 raw.jsonl；失败记录见 failures.jsonl。",
        "断点续跑：直接重跑本脚本即可，已完成的池会自动跳过。",
        "=" * 50,
        "安全：未写生产库、未发商品变更事件、未调用线上同步接口、未借用线上缓存通行证。",
    ]
    (OUT / "结果说明.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")
    log(f"完成：goods.csv {n} 行 / {len(CSV_FIELDS)} 列，说明见 output/结果说明.txt")


if __name__ == "__main__":
    main()
