#!/usr/bin/env python3
"""周频流水线编排（用户 2026-11-23 拍板：全链路周频自动化）。

阶段（用法：python3 weekly_pipeline.py <阶段>）：
  sunday  周日 20:00  — 京东VOP全量刷新 + enrich本网ID + 猫超list全量 + 新品详情回查（增量，只补list里有而详情里没有的品）
  price   接在 sunday 后 — 猫超全量价格刷新（约10小时，断点续跑，后台运行）
  monday  周一 06:30  — 等 sunday 完成 → 本网重筛 → 货盘一定价 → 猫超宽表 → 货盘二 → 自动归档进 交付物/2026-Wxx/ → 登记清单 → 提交推送

衔接机制：weekly_work/ 下的哨兵文件（sunday.done / price.done）。sunday/price 用 nohup 方式后台启动后立即返回，
monday 轮询哨兵。价格没跑完时不阻塞主链（回退用上周价格文件，宽表/聚合表照常出，只是价格旧一周）。

所有路径 __file__ 推导（R1），子脚本的环境变量覆盖全部保留。
"""
import datetime
import hashlib
import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PY = os.environ.get("WEEKLY_PY", str(Path.home() / "ai-workspace/local/venv/bin/python"))
PRICING = ROOT / "pricing_output"
ALLPOOL = ROOT / "allpool_output"
SUPPLY = ROOT / "supply_pull"
WORK = ROOT / "weekly_work"
TODAY = datetime.datetime.now().strftime("%Y-%m-%d")
WEEK = datetime.datetime.now().strftime("%G-W%V")

LOG = WORK / f"weekly_{TODAY}.log"


def log(msg):
    line = f"[{datetime.datetime.now():%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(script, cwd, desc, env_extra=None, sentinel=None, args_extra=None):
    """前台跑一个子脚本，成功写哨兵；失败抛异常中断流水线。"""
    env = {**os.environ, **(env_extra or {})}
    cmd = [PY, str(script), *(args_extra or [])]
    if os.environ.get("WEEKLY_DRYRUN"):
        log(f"▶ [演练] {desc}: {' '.join(cmd[1:])}  (cwd={cwd.name})")
        if sentinel:
            sentinel.write_text("dryrun", encoding="utf-8")
        return
    log(f"▶ {desc}: {script.name}")
    r = subprocess.run(cmd, cwd=str(cwd), env=env)
    if r.returncode != 0:
        raise RuntimeError(f"{desc} 失败（exit {r.returncode}）：{script.name}")
    if sentinel:
        sentinel.write_text(datetime.datetime.now().isoformat(), encoding="utf-8")
    log(f"✓ {desc} 完成")


def spawn_detached(script, cwd, desc, env_extra=None, args_extra=None):
    """后台启动（nohup 等价：脱离会话进程组），立即返回。完成状态靠哨兵。"""
    env = {**os.environ, **(env_extra or {})}
    log(f"🌙 后台启动 {desc}: {script.name}")
    lf = open(WORK / f"{desc}.log", "ab")
    subprocess.Popen([PY, str(script), *(args_extra or [])], cwd=str(cwd), env=env,
                     stdout=lf, stderr=subprocess.STDOUT,
                     start_new_session=True)
    log(f"   已脱离会话运行，日志 weekly_work/{desc}.log")


def wait_sentinel(sentinel, timeout_hours, label):
    deadline = time.time() + timeout_hours * 3600
    while time.time() < deadline:
        if sentinel.exists():
            log(f"✓ 哨兵到达：{label}")
            return True
        time.sleep(120)
    log(f"⚠ 等待超时（{timeout_hours}h）：{label} 未完成")
    return False


# ---------------- sunday ----------------

def _clear_sentinels():
    for s in ("sunday.done", "price.done"):
        p = WORK / s
        if p.exists():
            p.unlink()


def phase_sunday():
    WORK.mkdir(exist_ok=True)
    _clear_sentinels()
    run(SUPPLY / "jdvop_pull.py", SUPPLY, "京东VOP全量拉取",
        sentinel=WORK / "sunday_vop.done")
    run(SUPPLY / "enrich_bl_ids.py", SUPPLY, "京东VOP本网ID映射",
        sentinel=WORK / "sunday_enrich.done")
    run(SUPPLY / "maochao_list_10k.py", SUPPLY, "猫超list全量",
        args_extra=["200000"], sentinel=WORK / "sunday_list.done")
    # 详情：断点=所有历史详情并集，本周只补新品
    run(SUPPLY / "maochao_full_detail.py", SUPPLY, "猫超新品详情回查(增量)",
        env_extra={"MAOCHAO_WORKERS": "3"}, sentinel=WORK / "sunday.done")


def _price_worker():
    """价格后台工人：跑完写 price.done 哨兵（内部命令，勿直接调用）。"""
    r = subprocess.run([PY, str(SUPPLY / "maochao_price.py")], cwd=str(SUPPLY),
                       env={**os.environ, "MAOCHAO_WORKERS": "5", "MAOCHAO_SLEEP": "0.05"})
    if r.returncode == 0:
        (WORK / "price.done").write_text(datetime.datetime.now().isoformat(), encoding="utf-8")


def phase_price():
    WORK.mkdir(exist_ok=True)
    if not (WORK / "sunday.done").exists():
        log("⚠ 无 sunday.done，仍启动价格刷新（输入会取最新list）")
    spawn_detached(ROOT / "weekly_pipeline.py", ROOT, "price", args_extra=["_price_worker"])
    log("价格后台任务已启动（约10h，断点续跑）；monday 阶段轮询 price.done。")


# ---------------- monday ----------------

def phase_monday():
    WORK.mkdir(exist_ok=True)
    if os.environ.get("WEEKLY_DRYRUN"):
        for s in ("sunday.done", "price.done"):
            (WORK / s).write_text("dryrun", encoding="utf-8")
    # 1. 等 sunday（最多等 3 小时；超时则报错退出，人工介入）
    if not wait_sentinel(WORK / "sunday.done", 3, "sunday 外部货源阶段"):
        raise RuntimeError("sunday 阶段未完成，中断。请检查 weekly_work/ 下日志。")
    price_ready = wait_sentinel(WORK / "price.done", 1, "猫超价格(可选)")

    # 2. 本网池重筛五步
    run(PRICING / "pull_pool_sites.py", PRICING, "① 四站上架关系")
    run(PRICING / "aggregate_pool_candidates.py", PRICING, "② 四站全上架聚合")
    run(PRICING / "filter_stock.py", PRICING, "③ 四站库存筛选")
    run(PRICING / "pull_standard_pool_costs.py", PRICING, "④ ERP成本明细")
    run(PRICING / "pull_pool_product_info.py", PRICING, "⑤ 商品信息")

    # 3. 选品v2（温控剔除）
    run(ALLPOOL / "rescreen_pool.py", ALLPOOL, "⑥ 选品v2(冷链常温剔除)")

    # 4. 货盘一：京东相似匹配 + 定价
    run(PRICING / "match_pool_vop.py", PRICING, "⑦ 本网×京东VOP相似匹配")
    run(PRICING / "build_standard_pool_pricing.py", PRICING, "⑧ 货盘一定价",
        env_extra={
            "PRICING_POOL": str(ALLPOOL / f"标准池_选品v2_{TODAY}.jsonl"),
            "PRICING_INFO": str(PRICING / f"源数据_标准池商品信息_{TODAY}.jsonl"),
            "PRICING_COST": str(PRICING / f"源数据_标准池ERP站点成本_{TODAY}.jsonl"),
            "PRICING_MATCH": str(PRICING / f"标准池京东相似匹配_{TODAY}.jsonl"),
        })

    # 5. 猫超宽表（价格旧一周也不阻塞，脚本自动取最新）
    if not price_ready:
        log("⚠ 猫超价格未就绪，宽表使用上周价格文件（本周聚合表价格旧一周）")
    run(SUPPLY / "build_maochao_wide.py", SUPPLY, "⑨ 猫超SKU宽表")

    # 6. 货盘二链路
    run(ALLPOOL / "extract_maochao_barcodes.py", ALLPOOL, "⑩ 猫超条码索引")
    run(ALLPOOL / "pull_bl_barcodes.py", ALLPOOL, "⑪ 本网条码")
    run(ALLPOOL / "match_barcodes.py", ALLPOOL, "⑫ 条码对齐")
    run(ALLPOOL / "match_bl_maochao.py", ALLPOOL, "⑬ 本网×猫超相似")
    run(ALLPOOL / "task1_maochao_price.py", ALLPOOL, "⑭ 任务1猫超折算补价")
    run(ALLPOOL / "match_jd_allpool.py", ALLPOOL, "⑮ 京东×本网匹配", args_extra=["bl"])
    run(ALLPOOL / "match_jd_allpool.py", ALLPOOL, "⑯ 京东×猫超匹配", args_extra=["mc"])
    run(ALLPOOL / "build_allpool.py", ALLPOOL, "⑰ 全货盘聚合表")

    # 7. 归档
    if os.environ.get("WEEKLY_DRYRUN"):
        log("[演练] 跳过 git 归档与提交")
        return
    archive()
    log("🎉 周频流水线完成")


# ---------------- archive ----------------

def _upsert_inventory(moved):
    """交付文件清单.csv 按路径 upsert（R4/R6）。"""
    inv = ROOT / "00_项目汇总" / "交付文件清单.csv"
    rows = {}
    if inv.exists():
        for l in inv.read_text(encoding="utf-8-sig").splitlines():
            if l.strip():
                rows[l.split(",", 1)[0]] = l
    header = "相对路径,文件名,字节数,修改时间,SHA256"
    out = [rows.pop(header.split(",")[0], header)]
    for src, dst in moved:
        rel = str(dst.relative_to(ROOT))
        rows.pop(rel, None)
        size = dst.stat().st_size
        mtime = datetime.datetime.fromtimestamp(dst.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        sha = hashlib.sha256(dst.read_bytes()).hexdigest()
        rows[rel] = f"{rel},{dst.name},{size},{mtime},{sha}"
    inv.write_text("\n".join(out + list(rows.values())) + "\n", encoding="utf-8-sig")


def archive():
    """把本周产出移入 交付物/2026-Wxx/，登记清单，提交推送。"""
    wdir = ROOT / "交付物" / WEEK
    (wdir / "01_货盘一_本来定制").mkdir(parents=True, exist_ok=True)
    (wdir / "02_货盘二_技术公司全货盘").mkdir(parents=True, exist_ok=True)
    (wdir / "03_底稿数据").mkdir(parents=True, exist_ok=True)
    log(f"归档到 {wdir.relative_to(ROOT)}")

    plan = [
        (PRICING / f"分销标准池定价_{TODAY}.xlsx", wdir / "01_货盘一_本来定制"),
        (PRICING / f"定价统计_{TODAY}.json", wdir / "01_货盘一_本来定制"),
        (ALLPOOL / f"全货盘聚合表_{TODAY}.csv", wdir / "02_货盘二_技术公司全货盘"),
        # 本网底稿
        (PRICING / f"源数据_分销池四站上架关系_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (PRICING / f"源数据_标准池ERP站点成本_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (PRICING / f"源数据_标准池商品信息_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (PRICING / f"标准池_四站全上架有库存_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (PRICING / f"标准池候选_四站全上架_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (PRICING / f"标准池京东相似匹配_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (ALLPOOL / f"标准池_选品v2_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (ALLPOOL / f"剔除清单_选品v2_{TODAY}.jsonl", wdir / "03_底稿数据"),
        # 货盘二底稿（对齐文件在 allpool_output，文件名稳定）
        (ALLPOOL / "对齐_本网猫超.jsonl", wdir / "03_底稿数据"),
        (ALLPOOL / "对齐_京东_bl.jsonl", wdir / "03_底稿数据"),
        (ALLPOOL / "对齐_京东_mc.jsonl", wdir / "03_底稿数据"),
        (ALLPOOL / "任务1_猫超市场价.jsonl", wdir / "03_底稿数据"),
        (ALLPOOL / "本网条码.jsonl", wdir / "03_底稿数据"),
        (ALLPOOL / "猫超条码索引.jsonl", wdir / "03_底稿数据"),
        # 外部货源底稿
        (SUPPLY / "output" / "本来生活VOP_goods_本网ID.csv", wdir / "03_底稿数据"),
        (SUPPLY / "output" / f"猫超商品宽表SKU级_{TODAY}.csv", wdir / "03_底稿数据"),
        (SUPPLY / "output" / f"猫超产品_list全量147363个_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (SUPPLY / "output" / f"猫超详情_全量_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (SUPPLY / "output" / f"猫超SKU_全量_{TODAY}.jsonl", wdir / "03_底稿数据"),
        (SUPPLY / "output" / f"猫超价格_全量_{TODAY}.jsonl", wdir / "03_底稿数据"),
        # 取数说明（溯源凭证）
        (PRICING / f"取数说明_分销池四站上架关系_{TODAY}.json", wdir / "03_底稿数据"),
        (PRICING / f"取数说明_标准池ERP站点成本_{TODAY}.json", wdir / "03_底稿数据"),
        (PRICING / f"取数说明_标准池商品信息_{TODAY}.json", wdir / "03_底稿数据"),
        (PRICING / f"库存筛选说明_{TODAY}.json", wdir / "03_底稿数据"),
        (PRICING / f"聚合说明_标准池候选_{TODAY}.json", wdir / "03_底稿数据"),
    ]
    moved, missing = [], []
    for src, dst_dir in plan:
        if not src.exists():
            missing.append(src.name)
            continue
        dst = dst_dir / src.name
        r = subprocess.run(["git", "mv", str(src.relative_to(ROOT)), str(dst.relative_to(ROOT))],
                           cwd=str(ROOT), capture_output=True)
        if r.returncode != 0:  # 未跟踪/被忽略的文件 git mv 会失败，退回普通移动
            src.rename(dst)
        moved.append((src, dst))
    if missing:
        log(f"⚠ 归档缺文件（未产出？）：{'、'.join(missing)}")
    _upsert_inventory(moved)
    log(f"归档 {len(moved)} 个文件")

    subprocess.run(["git", "add", "-A"], cwd=str(ROOT), check=True)
    msg = f"weekly: {WEEK} 周频产出归档（货盘一+货盘二+底稿，流水线自动提交）"
    subprocess.run(["git", "commit", "-m", msg], cwd=str(ROOT), check=True)
    r = subprocess.run(["git", "push", "origin", "main"], cwd=str(ROOT),
                       capture_output=True, text=True)
    if r.returncode != 0:
        log(f"⚠ 推送失败（尚未同步）：{r.stderr[-200:]}")
    else:
        head = subprocess.run(["git", "log", "--oneline", "-1"], cwd=str(ROOT),
                              capture_output=True, text=True).stdout.strip()
        log(f"✓ 已推送 {head}")


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in ("sunday", "price", "monday", "_price_worker"):
        print("用法: python3 weekly_pipeline.py sunday|price|monday")
        sys.exit(1)
    if sys.argv[1] == "_price_worker":
        WORK.mkdir(exist_ok=True)
        _price_worker()
    else:
        {"sunday": phase_sunday, "price": phase_price, "monday": phase_monday}[sys.argv[1]]()
