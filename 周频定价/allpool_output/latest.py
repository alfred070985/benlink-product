#!/usr/bin/env python3
"""周频流水线共享的路径解析：环境变量优先，否则全仓按文件名日期取最新（R1 自包含）。

位置：周频定价/allpool_output/latest.py → 仓库根 = parents[2]。
用 rglob 按文件名全仓搜索（工作目录 + 周频定价/交付物 归档），目录结构再调整也不受影响。
"""
import os
import re
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def _date_key(p):
    """排序键：文件名里的日期优先，其次修改时间。避免工作区旧文件影子归档新文件。"""
    m = re.search(r"(20\d{2}-\d{2}-\d{2})", p.name)
    if m:
        return m.group(1)
    return datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d")


def _cands(fname):
    return [p for p in ROOT.rglob(fname) if p.is_file()]


def latest(rel_glob, env=None):
    """按文件名(支持*通配)找全仓最新文件；env 指定的环境变量可覆盖。找不到返回 None。"""
    if env and os.environ.get(env):
        return Path(os.environ[env])
    fname = rel_glob.split("/")[-1]
    cands = sorted(_cands(fname), key=_date_key)
    return cands[-1] if cands else None


def merged_lines(rel_glob, env=None):
    """跨所有历史文件按 product_id 去重（后写的覆盖先写的），返回去重后的行列表。
    周频增量下详情散落在多个日期文件里，消费方需要全量视图时用本函数。"""
    import json as _json
    if env and os.environ.get(env):
        return [l for l in open(os.environ[env], encoding="utf-8") if l.strip()]
    fname = rel_glob.split("/")[-1]
    cands = sorted(_cands(fname), key=_date_key)
    seen = {}
    for fp in cands:
        for l in open(fp, encoding="utf-8"):
            if not l.strip():
                continue
            try:
                k = _json.loads(l)["product_id"]
            except Exception:
                continue
            seen[k] = l
    return list(seen.values())
