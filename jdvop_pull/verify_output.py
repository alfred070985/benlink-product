#!/usr/bin/env python3
"""全量拉取结果对账：行数一致性、重复判定、按池分布、字段填充率，并生成结果说明.txt。"""
import csv
import json
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

OUT = Path("/Users/alfredmini/benlink_product/jdvop_pull/output")
CSV_PATH = OUT / "本来生活VOP_goods.csv"
RAW_PATH = OUT / "本来生活VOP_raw.jsonl"
FAIL_PATH = OUT / "本来生活VOP_failures.jsonl"
PROGRESS_PATH = OUT / "本来生活VOP_progress.json"
SUMMARY_PATH = OUT / "summary.json"

with CSV_PATH.open(encoding="utf-8-sig") as f:
    reader = csv.DictReader(f)
    columns = reader.fieldnames or []
    rows = list(reader)

def count_lines(path: Path) -> int:
    """逐行流式计数，避免把数百 MB 的原始记录一次性读进内存。"""
    if not path.exists():
        return 0
    with path.open(encoding="utf-8") as f:
        return sum(1 for line in f if line.strip())


raw_line_count = count_lines(RAW_PATH)
fail_line_count = count_lines(FAIL_PATH)
progress = json.loads(PROGRESS_PATH.read_text(encoding="utf-8")) if PROGRESS_PATH.exists() else {}
summary = json.loads(SUMMARY_PATH.read_text(encoding="utf-8")) if SUMMARY_PATH.exists() else {}

# 一张表按「商品池 + SKU」建记录：同池内重复才算错误；同一 SKU 落在多个池是京东池配置的正常重叠
row_keys = [(r["pool_id"], r["sku_id"]) for r in rows]
unique = len(set(r["sku_id"] for r in rows))
duplicate_rows = len(row_keys) - len(set(row_keys))
sku_in_pools = defaultdict(set)
for r in rows:
    sku_in_pools[r["sku_id"]].add(r["pool_id"])
cross_pool_skus = sum(1 for pools in sku_in_pools.values() if len(pools) > 1)

by_pool = Counter(r["pool_id"] for r in rows)
key_fields = ["sku_name", "brand_name", "category", "sale_price", "jd_price",
              "sale_state", "sku_state", "image_primary", "is_can_vat"]
fill = {k: sum(1 for r in rows if str(r.get(k, "")).strip()) for k in key_fields}
fill_rate = {k: (f"{v / len(rows) * 100:.1f}%" if rows else "0%") for k, v in fill.items()}

report = {
    "verified_at": datetime.now().isoformat(timespec="seconds"),
    "csv_rows": len(rows),
    "csv_columns": len(columns),
    "columns": columns,
    "unique_sku_ids": unique,
    "duplicate_rows": duplicate_rows,
    "cross_pool_skus": cross_pool_skus,
    "raw_jsonl_lines": raw_line_count,
    "csv_vs_raw_consistent": len(rows) == raw_line_count,
    "failure_lines": fail_line_count,
    "by_pool": dict(sorted(by_pool.items())),
    "pools_finished": sum(1 for v in progress.values() if v.get("done")),
    "pools_total": len(progress),
    "field_fill": fill,
    "field_fill_rate": fill_rate,
    "accounts": summary.get("accounts", []),
    "token_rule": summary.get("token_rule", ""),
    "side_effects": summary.get("side_effects", ""),
}
(OUT / "verify.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

skipped = [a.get("account") for a in summary.get("accounts", []) if a.get("status") != "ok"]
ok_accounts = [a.get("account") for a in summary.get("accounts", []) if a.get("status") == "ok"] or ["本来生活VOP"]
lines = [
    "京东 VOP 全量商品拉取 · 结果说明",
    f"生成时间：{report['verified_at']}",
    "",
    "一、数据口径",
    f"- 账号：{', '.join(ok_accounts)}（线上缓存中实际生效的账号）",
    f"- 商品池：本网配置表 third_code=4 / JDGoodsPoolIdList 中的 {len(progress)} 个池",
    "- 数据来源：京东 VOP 实时接口（商品池分页 → 商品详情 / 售卖价 / 图片 / 可售性 / 上下架状态）",
    "- 分页口径：与 tvopx-servicehost 的 JDVopGoodsSyncService 一致，每页 100 条",
    "",
    "二、Token 与生产安全",
    "- token 仅从本来生活 online 的 Redis 缓存读取（合规只读查询平台 GET），运行期间每 25 页重读一次",
    "- 全程未调用京东授权/刷新接口，未生成任何新 token",
    "- 未写生产库、未发布商品变更事件、未调用线上 /goods/sync/all 或 /goods/sync/sku",
    "- 未修改线上代码、未发布服务",
    "",
    "三、产物",
    f"- 本来生活VOP_goods.csv：表格，{len(rows)} 行 / {len(columns)} 列（Excel 可直接打开，UTF-8 BOM）",
    f"- 本来生活VOP_raw.jsonl：逐条原始记录，{raw_line_count} 行，保留每个接口的完整返回",
    f"- 本来生活VOP_failures.jsonl：失败清单，{fail_line_count} 条",
    f"- 本来生活VOP_progress.json：断点续跑进度，{report['pools_finished']}/{report['pools_total']} 个池已完成",
    "- verify.json：机器可读对账结果；summary.json：本次运行账号级汇总",
    "",
    "四、对账",
    f"- 记录行数：{len(rows)}；去重后唯一 SKU：{unique}",
    f"- 同一商品池内重复 SKU：{duplicate_rows} 行（应为 0）",
    f"- 跨池重复 SKU：{cross_pool_skus} 个（同一商品被配置在多个池，属正常业务现象，每池各保留一条记录）",
    f"- CSV 与原始记录行数一致：{report['csv_vs_raw_consistent']}",
    f"- 各池条数：{', '.join(f'{k}={v}' for k, v in report['by_pool'].items())}",
    f"- 关键字段有值行数：{', '.join(f'{k}={v}' for k, v in fill.items())}",
    f"- 价格覆盖率：{fill_rate.get('sale_price', '-')}（不可售商品京东不返回价格，留空属正常）",
    f"- 失败记录：{fail_line_count} 条",
    "",
    "五、未完成项",
]
if skipped:
    lines.append(f"- 账号 {', '.join(str(s) for s in skipped)}：线上缓存中无有效 token，按「不新生成 token」红线跳过，未做任何拉取。")
    lines.append("- 如需补齐该账号数据，只能由业务方在 online 正常授权流程中产生缓存 token 后再续跑，本工具不会代为申请。")
else:
    lines.append("- 无")
lines += [
    "",
    "六、复跑方式",
    "使用本机 Python 环境执行 jdvop_pull.py（默认全量，已有进度自动续跑）；--sample 为小样本试跑。",
    "工具内所有凭据从本机受限临时目录读取，代码中不含任何密钥。",
]
(OUT / "结果说明.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")

print(json.dumps({k: v for k, v in report.items() if k != "columns"}, ensure_ascii=False, indent=2))
print(f"已生成：{OUT / 'verify.json'}、{OUT / '结果说明.txt'}")
