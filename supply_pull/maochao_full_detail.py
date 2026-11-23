#!/usr/bin/env python3
"""猫超全量产品详情+SKU 慢速回查（夜间任务，断点续跑）。

- 输入：猫超产品_list全量*.jsonl（14.7 万 product_id）
- 逐个调 alibaba.ascp.channel.distributor.product.detail（含 SKU）
- 产物：
    猫超详情_全量_<日期>.jsonl   每行 {"product_id":..., "data":...详情全字段}
    猫超SKU_全量_<日期>.jsonl    每行一条 SKU（多规格品展开）
    猫超详情_失败_<日期>.jsonl   回查失败的 id 与原因
- 断点续跑：重启时扫描已有产物，跳过已成功的 product_id。
- 限速：默认每请求 sleep 0.25s（约 10 小时跑完 14.7 万），可用 MAOCHAO_SLEEP 覆盖。

用法：python maochao_full_detail.py
"""
import glob
import json
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from maochao_pull import load_creds, GATEWAY, CHANNEL, SUB_CHANNEL
import hashlib
from urllib.parse import quote
import urllib.request

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
SLEEP = float(os.environ.get("MAOCHAO_SLEEP", "0.1"))
WORKERS = int(os.environ.get("MAOCHAO_WORKERS", "3"))
METHOD = "alibaba.ascp.channel.distributor.product.detail"
_lock = threading.Lock()


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
    if isinstance(v, dict):
        return v.get("string") or []
    return v or []


def skus_of(d):
    psd = d.get("product_sku_detail") or {}
    return psd.get("product_sku_detail_d_t_o") or [] if isinstance(psd, dict) else []


def main():
    list_files = sorted(glob.glob(str(OUT_DIR / "猫超产品_list全量*.jsonl")))
    if not list_files:
        raise SystemExit("找不到全量列表文件")
    ids = []
    for l in open(list_files[-1], encoding="utf-8"):
        if l.strip():
            ids.append(json.loads(l)["product_id"])
    print(f"全量产品 {len(ids)} 个", flush=True)

    today = datetime.now().strftime("%Y-%m-%d")
    detail_path = OUT_DIR / f"猫超详情_全量_{today}.jsonl"
    sku_path = OUT_DIR / f"猫超SKU_全量_{today}.jsonl"
    fail_path = OUT_DIR / f"猫超详情_失败_{today}.jsonl"

    # 断点：已成功的 id = 跨所有历史详情文件的并集（周频增量：老品跳过，只补新品）
    done = set()
    for fp in sorted(glob.glob(str(OUT_DIR / "猫超详情_全量_*.jsonl"))):
        for l in open(fp, encoding="utf-8"):
            if l.strip():
                done.add(json.loads(l)["product_id"])
    todo = [i for i in ids if i not in done]
    print(f"已完成(跨历史文件) {len(done)}，待回查 {len(todo)}", flush=True)

    creds = load_creds()
    t0 = time.time()
    cnt = {"ok": 0, "fail": 0}
    fd = detail_path.open("a", encoding="utf-8")
    fs = sku_path.open("a", encoding="utf-8")
    ff = fail_path.open("a", encoding="utf-8")

    def work(pid):
        biz = {"product_id": pid, "channel_code": CHANNEL,
               "sub_channel_code": SUB_CHANNEL, "include_sku": True}
        data, err = None, ""
        for attempt in range(8):
            try:
                resp = top_call(creds, biz)
                if "error_response" in resp:
                    er = resp["error_response"]
                    err = f"{er.get('code')}:{er.get('sub_code')}:{str(er.get('sub_msg',''))[:60]}"
                    # 已取消铺货授权 = 业务性失败，重试无用
                    if "NOT_AUTH" in str(er.get("sub_code")):
                        break
                    time.sleep(min(2 * (attempt + 1), 15))
                    continue
                keys = [k for k in resp if k.endswith("_response")]
                result = resp[keys[0]].get("result", {}) if keys else {}
                if result.get("result_success") and result.get("data"):
                    data = result["data"]
                    break
                err = json.dumps(result, ensure_ascii=False)[:200] or "empty"
            except Exception as exc:
                err = str(exc)[:200]
            time.sleep(min(2 * (attempt + 1), 15))
        with _lock:
            if data:
                fd.write(json.dumps({"product_id": pid, "data": data}, ensure_ascii=False) + "\n")
                for s in skus_of(data):
                    fs.write(json.dumps({"product_id": pid, **s}, ensure_ascii=False) + "\n")
                cnt["ok"] += 1
            else:
                ff.write(json.dumps({"product_id": pid, "error": err}, ensure_ascii=False) + "\n")
                cnt["fail"] += 1
            i = cnt["ok"] + cnt["fail"]
            if i % 2000 == 0:
                fd.flush(); fs.flush(); ff.flush()
                el = time.time() - t0
                rate = i / el
                eta = (len(todo) - i) / rate if rate else 0
                print(f"[{datetime.now():%H:%M:%S}] {i}/{len(todo)} 成功{cnt['ok']} 失败{cnt['fail']} "
                      f"{rate:.1f}/s ETA {eta/3600:.1f}h", flush=True)
        time.sleep(SLEEP)

    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        list(ex.map(work, todo))
    fd.close(); fs.close(); ff.close()
    print(json.dumps({"回查完成": cnt["ok"], "失败": cnt["fail"], "总数": len(ids)},
                     ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
