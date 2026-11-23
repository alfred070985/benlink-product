#!/usr/bin/env python3
"""交付清单对账（MULTI_MACHINE_RULES R6）。

只读校验：把仓库内实际交付物与 00_项目汇总/交付文件清单.csv 逐项对账，
检查「清单有而仓库无」「仓库有而清单无」「字节数不符」「SHA256 不符」「路径重复」「格式错误」。

用法：python3 05_数据与脚本/处理脚本/verify_delivery_inventory.py
退出码：0=全部一致；1=存在差异；2=清单格式无法读取。

本脚本不修改任何文件。发现差异时，按 R5 以相对路径为主键 upsert 修正清单，
或在交付物创建当轮（R4）登记，再重跑本脚本。
"""

import csv
import hashlib
import io
import sys
from pathlib import Path

# R1：仓库根由 __file__ 推导，不写死本机绝对路径
REPO = Path(__file__).resolve().parents[2]
INVENTORY = REPO / "00_项目汇总" / "交付文件清单.csv"

# R4 交付物口径：六个业务目录下的全部文件 + 交付汇总文档。
# 治理类文件（.ai/、根目录说明、交付清单自身、新电脑接手指引等）不属于交付物，不登记。
SCOPE_DIRS = (
    "01_源数据",
    "02_报价表",
    "03_入盘建议",
    "04_履约成本",
    "05_数据与脚本",
    "06_校验记录",
)
EXTRA_FILES = ("00_项目汇总/benlink-product_交付汇总.md",)
HEADER = ["相对路径", "文件名", "字节数", "修改时间", "SHA256"]


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def actual_deliverables():
    result = set()
    for top in SCOPE_DIRS:
        directory = REPO / top
        if not directory.is_dir():
            continue
        for p in directory.rglob("*"):
            if p.is_file() and ".git" not in p.parts:
                result.add(p.relative_to(REPO).as_posix())
    for rel in EXTRA_FILES:
        if (REPO / rel).is_file():
            result.add(rel)
    return result


def main() -> int:
    if not INVENTORY.is_file():
        print(f"[错误] 找不到交付清单：{INVENTORY}")
        return 2

    raw = INVENTORY.read_bytes()
    try:
        rows = list(csv.reader(io.StringIO(raw.decode("utf-8-sig"))))
    except UnicodeDecodeError as exc:
        print(f"[错误] 交付清单不是 UTF-8：{exc}")
        return 2

    problems = []
    if not rows or rows[0] != HEADER:
        problems.append(f"表头不符：实际 {rows[0] if rows else '空'}，期望 {HEADER}")

    data = rows[1:] if rows else []
    bad_shape = [i + 2 for i, r in enumerate(data) if len(r) != 5 or not all(r)]
    if bad_shape:
        problems.append(f"列数不全或存在空字段的行号：{bad_shape[:10]}")

    listed = {}
    duplicates = []
    for r in data:
        if len(r) != 5:
            continue
        rel = r[0]
        if rel in listed:
            duplicates.append(rel)
        listed[rel] = r

    if duplicates:
        problems.append(f"清单内路径重复 {len(duplicates)} 条：{duplicates[:10]}")

    actual = actual_deliverables()
    listed_paths = set(listed)
    missing_in_inventory = sorted(actual - listed_paths)
    # 清单登记了但按口径不算交付物的文件：文件仍在磁盘上，属于登记范围问题而非丢失
    out_of_scope = sorted(p for p in listed_paths - actual if (REPO / p).is_file())
    # 清单登记了且磁盘上确实没有：真正的文件丢失
    missing_on_disk = sorted(p for p in listed_paths - actual if not (REPO / p).is_file())

    if missing_in_inventory:
        problems.append(f"仓库有、清单未登记 {len(missing_in_inventory)} 个：{missing_in_inventory[:10]}")
    if missing_on_disk:
        problems.append(f"清单有、仓库无（文件丢失）{len(missing_on_disk)} 个：{missing_on_disk[:10]}")
    if out_of_scope:
        problems.append(f"清单登记了非交付物 {len(out_of_scope)} 个：{out_of_scope[:10]}")

    size_bad, hash_bad, unreadable = [], [], []
    for rel, r in listed.items():
        p = REPO / rel
        if not p.is_file():
            continue
        try:
            size = p.stat().st_size
            if r[2].strip() != str(size):
                size_bad.append(rel)
            if sha256(p).lower() != r[4].strip().lower():
                hash_bad.append(rel)
        except OSError as exc:
            unreadable.append(f"{rel}（{exc}）")
    if size_bad:
        problems.append(f"字节数不符 {len(size_bad)} 个：{size_bad[:10]}")
    if hash_bad:
        problems.append(f"SHA256 不符 {len(hash_bad)} 个：{hash_bad[:10]}")
    if unreadable:
        problems.append(f"无法读取 {len(unreadable)} 个：{unreadable[:5]}")

    print("=== 交付清单对账（R6 只读校验）===")
    print(f"  清单条目：{len(data)}    仓库交付物：{len(actual)}    路径重复：{len(duplicates)}")
    print(f"  漏登：{len(missing_in_inventory)}    丢失：{len(missing_on_disk)}    非交付物：{len(out_of_scope)}")
    print(f"  字节数不符：{len(size_bad)}    SHA256 不符：{len(hash_bad)}")
    if problems:
        print("\n结论：不一致")
        for item in problems:
            print(f"  - {item}")
        return 1
    print("\n结论：全部一致（清单与仓库交付物逐项对得上）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
