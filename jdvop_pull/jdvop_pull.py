#!/usr/bin/env python3
"""京东 VOP 全量商品拉取（本地只读工具）

依据：codeup tvopx.git / stable 分支 tvopx-servicehost 模块
  - JDVopLowLeveApi / JDVopSdk / JDVopGoodsSyncService 的接口与分页口径
  - JDVopTokenCache / CakeKey 的 Redis key 规则
  - JDVopConfigSupport 的商品池配置来源

铁律（写死在代码里，不可通过参数绕过）：
  1. token 只从线上 Redis 缓存读取（合规只读查询平台，GET 命令）；
     绝不调用京东授权/刷新接口，绝不调用 /jd-vop/oauth/*。
  2. 缓存取不到或失效 -> 立即停止，不申请新 token。
  3. 不写任何生产库、不发商品变更事件、不调用线上 sync/all 或 sync/sku。
  4. 所有凭据来自本机临时受限文件，不写死在代码、不打印。
"""
import argparse
import csv
import hashlib
import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import requests
import yaml

ROOT = Path("/Users/alfredmini/benlink_product/jdvop_pull")
CONF_DIR = Path("/tmp/benlai-online-conf")
APOLLO_YML = CONF_DIR / "application.yml"
ARCHERY_JSON = CONF_DIR / "archery.json"
ARCHERY_MODULE_DIR = "/Users/alfredmini/ai-workspace/shared/benlai-trace-debug/scripts/archery"
JD_URL = "https://api.jd.com/routerjson"
APOLLO_TOOLS = "http://apollo-tools-online.benlai.com"
TOKEN_KEY = "erp:third_vendor:jd:access_token:{}"
REDIS_INSTANCE = "233"          # [online]online2-cache.redis.rds.aliyuncs.com
MYSQL_INSTANCE = "147"          # [online]生产PolarDB-Newretail
POOL_DB = "benlai_third_vendor"
POOL_SQL = "select c_value from t_config where third_code = '4' and c_type = 'JDGoodsPoolIdList'"
PAGE_SIZE = 100                 # 与 JDVopGoodsSyncService 一致
TOKEN_REFRESH_PAGES = 25        # 每 25 页重读一次线上缓存（只读取，不刷新）

CSV_FIELDS = [
    "account", "pool_id", "sku_id", "sku_name", "brand_name", "category", "sale_unit",
    "sale_state", "sku_state", "is_can_vat", "ban_cause", "return_rule_type",
    "sale_price", "jd_price", "market_price", "naked_price", "tax_rate_percentage",
    "tax_code", "seo_model", "lowest_buy", "product_area", "color", "size", "weight",
    "length", "width", "height", "self_sell_type", "logistics_type", "image_primary",
    "image_count", "captured_at",
]


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def ensure_inputs():
    """确保线上配置与合规查询账号就绪；缺失时自动按 CI 方式补齐。"""
    CONF_DIR.mkdir(parents=True, exist_ok=True)
    os.chmod(CONF_DIR, 0o700)
    if not APOLLO_YML.exists():
        cicd = CONF_DIR / "AppCicd.yaml"
        cicd.write_text(
            "ENV: online\nAPP_NAME: tvopx-servicehost\nAPP_CONFIG:\n"
            "  application: default\n  application-files: default\n"
            "  application-certs: default\n  application-binaries: default\n",
            encoding="utf-8")
        with cicd.open("rb") as f:
            r = requests.post(f"{APOLLO_TOOLS}/upload", files={"file": f}, timeout=30)
        urls = [u.strip() for u in r.text.splitlines() if u.strip().endswith("application.yml")]
        if not urls:
            raise RuntimeError("无法从 Apollo 配置服务取得线上配置")
        APOLLO_YML.write_bytes(requests.get(urls[0], timeout=30).content)
        os.chmod(APOLLO_YML, 0o600)
        log("已从 Apollo 配置服务下载线上配置")
    if not ARCHERY_JSON.exists():
        sys.path.insert(0, "/Users/alfredmini/Benlink/ops-local/skills/benlai-trace-debug/scripts")
        import _common
        cfg = _common.load_config().get("archery", {})
        ARCHERY_JSON.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        os.chmod(ARCHERY_JSON, 0o600)
        log("已从本机加密配置解出只读查询账号")


def build_archery():
    sys.path.insert(0, ARCHERY_MODULE_DIR)
    import archery_client as ac
    client = ac.ArcheryClient(config_path=str(ARCHERY_JSON))
    if not client.login():
        raise RuntimeError("只读查询平台登录失败")
    return client


def archery_get(client, instance, db, command):
    """只发 GET / SELECT。平台侧另有命令白名单，写命令不可能通过。"""
    result = client.query(instance, str(db), command, limit_num=100)
    if result.get("is_masked"):
        raise RuntimeError("查询结果被平台脱敏，停止（不猜测 token）")
    rows = result.get("rows") or []
    cells = [str(c) for row in rows for c in row if c is not None and c != ""]
    return cells


def read_pool_ids(client):
    cells = archery_get(client, MYSQL_INSTANCE, POOL_DB, POOL_SQL)
    if not cells:
        raise RuntimeError("线上配置表未返回已配置商品池清单")
    pools = json.loads(cells[0])
    if not isinstance(pools, list) or not pools:
        raise RuntimeError("已配置商品池清单为空或格式异常")
    return [str(p) for p in pools]


def read_token(client, account):
    """唯一取 token 的入口：只读线上缓存。缺失返回 None，绝不生成。"""
    cells = archery_get(client, REDIS_INSTANCE, "0", f"get {TOKEN_KEY.format(account)}")
    value = cells[0] if cells else ""
    if not value or value.lower() in ("nil", "none", "(nil)"):
        return None
    return value


class JdVopClient:
    """按项目 JDVopSdk 的签名与参数规则直连京东 VOP 老接口。"""

    def __init__(self, app_info, token):
        self.app_key = app_info["app-key"]
        self.app_secret = app_info["app-secret"]
        self.token = token

    def call(self, method, biz, retries=3):
        last = None
        for attempt in range(1, retries + 1):
            try:
                params = {
                    "method": method,
                    "app_key": self.app_key,
                    "access_token": self.token,
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
                    "format": "json",
                    "v": "2.0",
                    "sign_method": "md5",
                    "360buy_param_json": json.dumps(biz, ensure_ascii=False, separators=(",", ":")),
                }
                src = self.app_secret + "".join(k + str(params[k]) for k in sorted(params)) + self.app_secret
                params["sign"] = hashlib.md5(src.encode("utf-8")).hexdigest().upper()
                resp = requests.post(JD_URL, data=params, timeout=30)
                resp.raise_for_status()
                payload = resp.json()
                key = method.replace(".", "_") + "_responce"
                rpc = payload[key]["openRpcResult"]
                if str(rpc.get("resultCode")) not in ("0000", "0") or rpc.get("success") is False:
                    raise RuntimeError(f"{method} resultCode={rpc.get('resultCode')} {rpc.get('resultMessage')}")
                return rpc.get("result")
            except Exception as exc:
                last = exc
                if attempt < retries:
                    time.sleep(0.6 * attempt)
        raise RuntimeError(f"{method} 重试 {retries} 次仍失败: {str(last)[:160]}")


def flatten(detail, price, image, sale, state):
    detail = detail or {}
    price = price or {}
    image = image or {}
    sale = sale or {}
    state = state or {}
    images = image.get("skuImageList") or []
    primary = ""
    for item in images:
        if item.get("isPrimary") == 1 or item.get("position") == 0:
            primary = item.get("shortPath") or item.get("imagePath") or ""
            break
    if not primary and images:
        primary = images[0].get("shortPath") or ""
    return {
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
    }


def load_progress(path):
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def save_progress(path, data):
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def run_account(account, app_info, pools, archery, out_dir, max_pages, workers, sample):
    token = read_token(archery, account)
    if not token:
        log(f"账号 {account}：线上缓存无有效 token -> 跳过（按铁律不生成新 token）")
        return {"account": account, "status": "skipped_no_cached_token", "skus": 0}
    jd = JdVopClient(app_info, token)
    log(f"账号 {account}：已取得线上缓存 token（长度 {len(token)}），开始拉取")

    safe = account.replace("/", "_")
    csv_path = out_dir / f"{safe}_goods.csv"
    raw_path = out_dir / f"{safe}_raw.jsonl"
    fail_path = out_dir / f"{safe}_failures.jsonl"
    progress_path = out_dir / f"{safe}_progress.json"
    write_header = not csv_path.exists()
    csv_f = csv_path.open("a", encoding="utf-8-sig", newline="")
    writer = csv.DictWriter(csv_f, fieldnames=CSV_FIELDS, extrasaction="ignore")
    if write_header:
        writer.writeheader()
    raw_f = raw_path.open("a", encoding="utf-8")
    fail_f = fail_path.open("a", encoding="utf-8")
    progress = load_progress(progress_path)
    total = 0
    failed = 0

    def record_failure(pool, sku, stage, err):
        nonlocal failed
        failed += 1
        fail_f.write(json.dumps({"pool_id": pool, "sku_id": sku, "stage": stage,
                                 "error": str(err)[:300], "ts": datetime.now().isoformat()},
                                ensure_ascii=False) + "\n")
        fail_f.flush()

    try:
        for pool in pools:
            state = progress.get(pool, {})
            if state.get("done"):
                log(f"池 {pool} 已完成，跳过")
                continue
            offset = state.get("offset", 0)
            pages = state.get("pages", 0)
            while True:
                if max_pages and pages >= max_pages:
                    log(f"池 {pool} 达到本次页数上限 {max_pages}")
                    break
                if pages and pages % TOKEN_REFRESH_PAGES == 0:
                    fresh = read_token(archery, account)
                    if not fresh:
                        raise RuntimeError("运行期间线上缓存 token 消失 -> 停止（不生成新 token）")
                    if fresh != jd.token:
                        jd.token = fresh
                        log("已重读线上缓存 token（线上服务刷新后的值）")
                try:
                    page = jd.call("jingdong.vop.goods.querySkuByPage",
                                   {"bizPoolId": pool, "offset": offset, "pageSize": PAGE_SIZE})
                except Exception as exc:
                    record_failure(pool, offset, "querySkuByPage", exc)
                    raise
                skus = page.get("skus") or []
                sku_csv = ",".join(str(s) for s in skus)
                sale_map, price_map, image_map, state_map = {}, {}, {}, {}
                if skus:
                    for method, target in [
                        ("jingdong.vop.goods.checkSkuSaleList", sale_map),
                        ("jingdong.vop.goods.getSkuImageList", image_map),
                        ("jingdong.vop.goods.getSkuStateList", state_map),
                    ]:
                        try:
                            rows = jd.call(method, {"skuId": sku_csv}) or []
                            for row in rows:
                                if row and row.get("skuId") is not None:
                                    target[int(row["skuId"])] = row
                        except Exception as exc:
                            record_failure(pool, "", method, exc)
                    can_sale = [s for s in skus if sale_map.get(int(s), {}).get("saleState") == 1]
                    if can_sale:
                        try:
                            rows = jd.call("jingdong.vop.goods.getSellPrice",
                                           {"skuId": ",".join(str(s) for s in can_sale)}) or []
                            for row in rows:
                                if row and row.get("skuId") is not None:
                                    price_map[int(row["skuId"])] = row
                        except Exception as exc:
                            record_failure(pool, "", "getSellPrice", exc)

                    def fetch_detail(sku):
                        try:
                            return sku, jd.call("jingdong.vop.goods.getSkuDetailInfo", {"skuId": int(sku)}), None
                        except Exception as exc:
                            return sku, None, exc

                    details = {}
                    with ThreadPoolExecutor(max_workers=workers) as pool_exec:
                        futures = [pool_exec.submit(fetch_detail, s) for s in skus]
                        for fut in as_completed(futures):
                            sku, detail, err = fut.result()
                            if err:
                                record_failure(pool, sku, "getSkuDetailInfo", err)
                            else:
                                details[int(sku)] = detail

                    captured = datetime.now().isoformat(timespec="seconds")
                    for sku in skus:
                        sku_id = int(sku)
                        detail = details.get(sku_id)
                        row = {"account": account, "pool_id": pool, "sku_id": sku_id,
                               "captured_at": captured}
                        row.update(flatten(detail, price_map.get(sku_id), image_map.get(sku_id),
                                           sale_map.get(sku_id), state_map.get(sku_id)))
                        writer.writerow(row)
                        raw_f.write(json.dumps({
                            "account": account, "pool_id": pool, "sku_id": sku_id,
                            "captured_at": captured,
                            "detail": detail,
                            "price": price_map.get(sku_id),
                            "image": image_map.get(sku_id),
                            "sale_state": sale_map.get(sku_id),
                            "sku_state": state_map.get(sku_id),
                        }, ensure_ascii=False) + "\n")
                        total += 1
                    csv_f.flush()
                    raw_f.flush()

                pages += 1
                next_offset = page.get("offset", offset + len(skus))
                remain = page.get("remainPage", 0)
                progress[pool] = {"offset": next_offset, "pages": pages,
                                  "done": remain == 0, "skus": progress.get(pool, {}).get("skus", 0) + len(skus)}
                save_progress(progress_path, progress)
                log(f"池 {pool} 第 {pages} 页：{len(skus)} 个 SKU，累计 {total}，剩余页 {remain}")
                if sample:
                    break
                if remain == 0:
                    break
                offset = next_offset
                time.sleep(0.12)
            if sample:
                break
    finally:
        csv_f.close()
        raw_f.close()
        fail_f.close()
    return {"account": account, "status": "ok", "skus": total, "failures": failed,
            "csv": str(csv_path), "raw": str(raw_path)}


def main():
    parser = argparse.ArgumentParser(description="京东 VOP 全量商品拉取（线上缓存 token 只读）")
    parser.add_argument("--sample", action="store_true", help="小样本试跑：第一个池 1 页")
    parser.add_argument("--max-pages", type=int, default=0, help="每个池最多拉取页数（0=不限）")
    parser.add_argument("--workers", type=int, default=8, help="商品详情并发数")
    args = parser.parse_args()

    ensure_inputs()
    out_dir = ROOT / ("output_sample" if args.sample else "output")
    out_dir.mkdir(parents=True, exist_ok=True)
    archery = build_archery()
    pools = read_pool_ids(archery)
    log(f"已配置商品池 {len(pools)} 个：{', '.join(pools)}")
    cfg = yaml.safe_load(APOLLO_YML.read_text(encoding="utf-8"))
    apps = cfg["jd-vop"]["sdk"]["apps"]
    log(f"线上账号 {len(apps)} 个，逐个校验线上缓存 token")

    results = []
    for app_info in apps:
        account = app_info["username"]
        try:
            results.append(run_account(account, app_info, pools, archery, out_dir,
                                       1 if args.sample else args.max_pages,
                                       args.workers, args.sample))
        except Exception as exc:
            log(f"账号 {account} 中止：{str(exc)[:200]}")
            results.append({"account": account, "status": "stopped", "error": str(exc)[:300]})
    summary = {"finished_at": datetime.now().isoformat(timespec="seconds"),
               "mode": "sample" if args.sample else "full",
               "pools": pools, "accounts": results,
               "token_rule": "只读线上 Redis 缓存；缺失即停止，未调用任何授权/刷新接口",
               "side_effects": "未写生产库、未发商品变更事件、未调用线上同步接口"}
    (out_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"完成，汇总写入 {out_dir / 'summary.json'}")
    for r in results:
        log(f"  {r['account']}: {r['status']} skus={r.get('skus', 0)} failures={r.get('failures', 0)}")


if __name__ == "__main__":
    main()
