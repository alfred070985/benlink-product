#!/usr/bin/env python3
"""猫超回查任务监控看门狗：每 10 分钟检查产出增长与进程状态，异常自动重启。

检查对象（supply_pull/output/）：
  猫超详情_全量_2026-11-20.jsonl + 失败文件  ← maochao_full_detail.py
  猫超价格_全量_2026-11-20.jsonl              ← maochao_price.py
判定：
  - 进程不在跑 且 行数未达基准 → 断点续跑重启
  - 连续两轮产出零增长 且 进程在跑 → 判定卡死，杀进程重启
输出：stdout 一行状态/轮；异常写 监控告警_<日期>.log
用法：python maochao_watchdog.py [间隔秒，默认600]
"""
import glob
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
VENV_PY = "/Users/alfredmini/ai-workspace/local/venv/bin/python"
BASELINE = 147363
INTERVAL = int(sys.argv[1]) if len(sys.argv) > 1 else 600

TASKS = [
    {"name": "详情", "proc": "maochao_full_detail.py", "files": ["猫超详情_全量_2026-11-20.jsonl", "猫超详情_失败_2026-11-20.jsonl"],
     "env": {"MAOCHAO_WORKERS": "2", "MAOCHAO_SLEEP": "0.3"}},
    {"name": "价格", "proc": "maochao_price.py", "files": ["猫超价格_全量_2026-11-20.jsonl", "猫超价格_失败_2026-11-20.jsonl"],
     "env": {"MAOCHAO_WORKERS": "5", "MAOCHAO_SLEEP": "0.05"}},
]


def total_lines(files):
    n = 0
    for fn in files:
        for cand in glob.glob(str(OUT_DIR / fn)):
            with open(cand, encoding="utf-8") as f:
                for _ in f:
                    n += 1
    return n


def proc_alive(pattern):
    r = subprocess.run(["pgrep", "-f", pattern], capture_output=True)
    return r.returncode == 0


def restart(task, reason):
    log = OUT_DIR / f"监控告警_{datetime.now():%Y-%m-%d}.log"
    msg = f"[{datetime.now():%H:%M:%S}] 重启{task['name']}任务：{reason}\n"
    with open(log, "a", encoding="utf-8") as f:
        f.write(msg)
    print(msg, end="", flush=True)
    env = {**os.environ, **task["env"]}
    subprocess.Popen([VENV_PY, str(HERE / task["proc"])],
                     cwd=str(HERE), env=env,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                     start_new_session=True)


def main():
    print(f"[{datetime.now():%H:%M:%S}] 监控启动，间隔 {INTERVAL}s", flush=True)
    prev = {}
    stall = {t["name"]: 0 for t in TASKS}
    while True:
        for t in TASKS:
            cur = total_lines(t["files"])
            last = prev.get(t["name"])
            prev[t["name"]] = cur
            alive = proc_alive(t["proc"])
            if last is None:
                print(f"[{datetime.now():%H:%M:%S}] {t['name']}: 基线 {cur} 行，进程{'在跑' if alive else '不在'}", flush=True)
                continue
            delta = cur - last
            remaining = max(BASELINE - cur, 0)
            rate = delta / (INTERVAL / 60)
            eta_h = remaining / (rate * 60) if rate > 0 else None
            status = "在跑" if alive else "已退出"
            eta_s = f"ETA {eta_h:.1f}h" if eta_h else ("完成/收敛" if remaining == 0 else "ETA ∞")
            print(f"[{datetime.now():%H:%M:%S}] {t['name']}: {cur} 行(+{delta}) {status} {rate:.0f}行/分 {eta_s}", flush=True)
            if not alive and cur < BASELINE * 0.98:
                restart(t, f"进程退出（{cur}/{BASELINE}），断点续跑")
                stall[t["name"]] = 0
            elif alive and delta <= 0 and cur < BASELINE * 0.98:
                stall[t["name"]] += 1
                if stall[t["name"]] >= 2:
                    subprocess.run(["pkill", "-f", t["proc"]])
                    time.sleep(2)
                    restart(t, f"连续两轮零产出（卡死嫌疑），强杀重启")
                    stall[t["name"]] = 0
            else:
                stall[t["name"]] = 0
        time.sleep(INTERVAL)


if __name__ == "__main__":
    main()
