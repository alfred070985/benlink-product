#!/usr/bin/env python3
"""从「天镜智价」用免费查询接口全量拉取价格历史 + 同品映射，落本地快照。

合规边界（用户2026-11-16确认）：
  - 只用页面正常的「查询」类接口，不调用导出的 mode=export_csv / export_cost，
    不伪造 confirm_cost，不绕过平台的积分计费逻辑 —— 全程不消耗积分。
  - 账号密码从本机凭证文件读取（不进 git），路径由 __file__ 推导，可用环境变量覆盖。

产出（均在 PRICING_DIR）：
  - 源数据_天镜价格历史全量_<日期>.jsonl   每行一条查询历史（每SKU保留最新一条由下游处理）
  - 源数据_天镜同品映射_<日期>.jsonl        全部已确认映射（P码<->SKU<->同款/不同规/相似）
  - 天镜取数日志_<日期>.json                 记录条数、耗时、时间范围、源文件指纹
"""
import http.cookiejar
import json
import os
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(os.environ.get("PRICING_DIR", HERE))
CRED = Path(os.environ.get(
    "TIANJING_CRED",
    Path.home() / "ai-workspace" / "local" / "tianjing_credentials.json",
))
TODAY = datetime.now().strftime("%Y-%m-%d")

HIST_OUT = PRICING_DIR / f"源数据_天镜价格历史全量_{TODAY}.jsonl"
MAP_OUT = PRICING_DIR / f"源数据_天镜同品映射_{TODAY}.jsonl"
LOG_OUT = PRICING_DIR / f"天镜取数日志_{TODAY}.json"

HIST_PAGE = 5000   # get_history 每批条数（实测稳定）
MAP_PAGE = 1000    # verified 列表每页条数


def build_opener():
    cred = json.loads(CRED.read_text(encoding="utf-8"))
    base = cred["base"].rstrip("/")
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    return opener, base, cred


def post_json(opener, base, path, payload, timeout=90):
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        base + path, data=data,
        headers={"Content-Type": "application/json", "Origin": base},
        method="POST",
    )
    with opener.open(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def get_json(opener, base, path, timeout=90):
    req = urllib.request.Request(base + path, headers={"Origin": base})
    with opener.open(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def login(opener, base, cred):
    res = post_json(opener, base, "/api/auth/login", {
        "employeeId": cred["employeeId"],
        "password": cred["password"],
        "remember": True,
    })
    if not res.get("success"):
        raise RuntimeError(f"登录失败: {res}")
    return res


def fetch_history(opener, base, out_path):
    """全量拉价格查询历史，按 id 倒序（最新在前），逐批 append。"""
    total_hint = None
    offset = 0
    written = 0
    t0 = time.time()
    with out_path.open("w", encoding="utf-8") as f:
        while True:
            res = post_json(opener, base, "/api/vop", {
                "action": "get_history",
                "payload": {"limit": HIST_PAGE, "offset": offset},
            })
            items = res.get("items", [])
            if total_hint is None:
                total_hint = int(res.get("total", 0))
            if not items:
                break
            for it in items:
                # 只留必要字段，避免 fullRequest/fullResponse 撑爆文件
                slim = {
                    "skuId": it.get("skuId"),
                    "jdPrice": it.get("jdPrice"),
                    "salePrice": it.get("salePrice"),
                    "status": it.get("status"),
                    "fetchTime": it.get("fetchTime") or it.get("_fetch_time"),
                }
                f.write(json.dumps(slim, ensure_ascii=False) + "\n")
            written += len(items)
            offset += len(items)
            if offset % 50000 < HIST_PAGE or written >= total_hint:
                print(f"  历史 {written}/{total_hint}  用时{time.time()-t0:.0f}s", flush=True)
            if written >= total_hint or len(items) < HIST_PAGE:
                break
    return written, total_hint, time.time() - t0


def fetch_mapping(opener, base, out_path):
    """全量拉已确认同品映射（免费列表接口，非导出）。"""
    page = 1
    written = 0
    total_hint = None
    t0 = time.time()
    with out_path.open("w", encoding="utf-8") as f:
        while True:
            res = get_json(opener, base, f"/api/admin/contribution/verified?page={page}&pageSize={MAP_PAGE}")
            items = res.get("items", [])
            if total_hint is None:
                total_hint = int(res.get("total", 0))
            if not items:
                break
            for it in items:
                f.write(json.dumps(it, ensure_ascii=False) + "\n")
            written += len(items)
            if written % 10000 < MAP_PAGE or written >= total_hint:
                print(f"  映射 {written}/{total_hint}  用时{time.time()-t0:.0f}s", flush=True)
            if written >= total_hint or len(items) < MAP_PAGE:
                break
            page += 1
    return written, total_hint, time.time() - t0


def main():
    print(f"[{TODAY}] 天镜智价免费接口取数开始（不消耗积分）", flush=True)
    opener, base, cred = build_opener()
    login_info = login(opener, base, cred)
    print(f"登录成功 role={login_info.get('role')}", flush=True)

    print("→ 拉取价格历史全量", flush=True)
    h_written, h_total, h_sec = fetch_history(opener, base, HIST_OUT)

    print("→ 拉取同品映射全量", flush=True)
    m_written, m_total, m_sec = fetch_mapping(opener, base, MAP_OUT)

    # 历史时间范围
    tmin = tmax = None
    with HIST_OUT.open(encoding="utf-8") as f:
        for line in f:
            o = json.loads(line)
            t = o.get("fetchTime") or o.get("_fetch_time")
            if t:
                if tmin is None or t < tmin: tmin = t
                if tmax is None or t > tmax: tmax = t

    log = {
        "取数时间": datetime.now().isoformat(timespec="seconds"),
        "站点": base,
        "账号": cred["employeeId"],
        "登录角色": login_info.get("role"),
        "方式": "免费查询接口(get_history + verified列表)，未调用导出、未消耗积分",
        "价格历史": {"文件": HIST_OUT.name, "条数": h_written, "接口报告总数": h_total,
                     "耗时秒": round(h_sec, 1), "时间范围": [tmin, tmax]},
        "同品映射": {"文件": MAP_OUT.name, "条数": m_written, "接口报告总数": m_total,
                     "耗时秒": round(m_sec, 1)},
    }
    LOG_OUT.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print("完成：", json.dumps(log, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
