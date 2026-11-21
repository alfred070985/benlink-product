#!/usr/bin/env python3
"""Offline continuation checks; never connects to Archery or production data."""
import argparse
import importlib.util
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PROJECT_ID = "benlink-product"
DEFAULT_ARCHERY_DIR = (
    Path.home()
    / "ai-workspace"
    / "shared"
    / "benlai-trace-debug"
    / "scripts"
    / "archery"
)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-archery", action="store_true")
    args = parser.parse_args()
    errors = []
    warnings = []

    required_files = [
        "START_HERE.md",
        "AGENTS.md",
        ".ai/project.yaml",
        ".ai/state.md",
        ".gitattributes",
        "00_项目汇总/benlink-product_交付汇总.md",
        "05_数据与脚本/查询脚本/统计PolarDB禄品平台商品.sql",
        "05_数据与脚本/处理脚本/archery_runtime.py",
    ]
    for relative in required_files:
        if not (ROOT / relative).is_file():
            errors.append(f"缺少必要文件：{relative}")

    identity = (ROOT / ".ai/project.yaml").read_text(encoding="utf-8")
    if f"project_id: {PROJECT_ID}" not in identity:
        errors.append("project_id 不匹配")

    script_results = []
    for path in sorted((ROOT / "05_数据与脚本").rglob("*.py")):
        relative = path.relative_to(ROOT).as_posix()
        try:
            compile(path.read_text(encoding="utf-8"), str(path), "exec")
            script_results.append({"file": relative, "syntax": "ok"})
        except Exception as exc:
            errors.append(f"Python 语法检查失败：{relative}: {exc}")

    attributes = (ROOT / ".gitattributes").read_text(encoding="utf-8")
    for pattern in ("*.xlsx", "*.jsonl"):
        if not any(line.startswith(pattern + " ") and "filter=lfs" in line
                   for line in attributes.splitlines()):
            errors.append(f"Git LFS 未跟踪：{pattern}")

    git_lfs = shutil.which("git-lfs") or shutil.which("git")
    if git_lfs is None:
        warnings.append("未找到 Git；首次推送前需安装 Git 和 Git LFS")
    else:
        result = subprocess.run(
            ["git", "lfs", "version"], capture_output=True, text=True, timeout=10
        )
        if result.returncode != 0:
            warnings.append("Git LFS 尚未安装或不可用")

    dependencies = {
        name: importlib.util.find_spec(name) is not None
        for name in ("openpyxl", "bs4")
    }
    for name, available in dependencies.items():
        if not available:
            warnings.append(f"当前 Python 未安装可选运行依赖：{name}")

    archery_dir = Path(
        os.environ.get("BENLAI_ARCHERY_DIR", str(DEFAULT_ARCHERY_DIR))
    ).expanduser()
    archery_ready = (archery_dir / "archery_client.py").is_file()
    if not archery_ready:
        message = "Archery 能力包未通过 BENLAI_ARCHERY_DIR 或统一软链接配置"
        if args.require_archery:
            errors.append(message)
        else:
            warnings.append(message)

    summary = {
        "project_id": PROJECT_ID,
        "status": "failed" if errors else "passed",
        "mode": "offline_no_production_connection",
        "required_files": "ok" if not any("必要文件" in x for x in errors) else "failed",
        "python_scripts_checked": len(script_results),
        "dependencies": dependencies,
        "archery_ready": archery_ready,
        "warnings": warnings,
        "errors": errors,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
