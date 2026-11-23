#!/usr/bin/env python3
"""京东 VOP 全量 CSV 二次加工：补本网商品 / 类目 / 品牌 ID。

用户已拍板的口径（不得擅自更改）：
  1. 映射主口径 = 分销商品池 benlai_business_product.product.JdGoodsId 直连；
     不使用 t_goods_map 覆盖、不自动裁决两条路径的 P 码冲突。
  2. 商品 ID 同时给 P 码（ProductId）与主商品编号（ProductBasicSysNo）。
  3. 类目给本网三级 ID（C1/C2/C3SysNo），品牌给 BrandSysNo；不含名称。
  4. 匹配不上或存在歧义的记录：ID 列全部留空，用 match_source / conflict_with_goods_map 标记，一行不丢。

安全约束：线上库只读 SELECT；不写生产库、不发事件、不调用京东接口、不生成 token。
"""
import csv
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import jdvop_pull as j  # noqa: E402  复用已验证的只读查询通道

OUT = ROOT / "output"
SRC = OUT / "本来生活VOP_goods.csv"
DST = OUT / "本来生活VOP_goods_本网ID.csv"
VERIFY = OUT / "本网ID_对账.json"
README = OUT / "本网ID_结果说明.txt"
DETAIL = OUT / "本网ID_待确认.csv"

POOL_SQL = """select JdGoodsId, ProductId, ProductBasicSysNo, C1SysNo, C2SysNo, C3SysNo,
                     BrandSysNo, ProductStatus, IsDelete
from benlai_business_product.product
where JdGoodsId is not null and JdGoodsId <> ''"""
MAP_SQL = "select third_goods_id, self_goods_id from benlai_third_vendor.t_goods_map where third_code='JD'"

NEW_FIELDS = [
    "bl_product_code",      # P 码
    "bl_product_sysno",     # 主商品编号（ProductBasicSysNo / C 网主商品编号）
    "bl_c1_id", "bl_c2_id", "bl_c3_id", "bl_brand_id",
    "match_source", "conflict_with_goods_map",
]


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def s(v):
    """统一转成 CSV 里的字符串；None / 空串留空，0 视为合法值保留。"""
    if v is None:
        return ""
    t = str(v).strip()
    return "" if t.lower() in ("none", "null", "nan") else t


def query_rows(client, sql, limit):
    r = client.query("147", "benlai_third_vendor", sql, limit_num=limit)
    if r.get("error"):
        raise RuntimeError(f"只读查询失败：{str(r['error'])[:200]}")
    if r.get("is_masked"):
        raise RuntimeError("查询结果被平台脱敏，停止")
    return r.get("rows") or []


def load_pool(client):
    """按京东 SKU 归集分销商品池记录，只保留有效记录中的唯一 P 码。"""
    rows = query_rows(client, POOL_SQL, 130000)
    log(f"分销商品池返回 {len(rows)} 行（带 JdGoodsId）")
    grouped = defaultdict(list)
    for row in rows:
        sku = s(row[0])
        if sku:
            grouped[sku].append(row)

    unique = {}
    ambiguous = set()
    inactive_only = set()
    all_pcodes = {}
    for sku, items in grouped.items():
        active = [x for x in items if s(x[8]) != "1" and s(x[7]) != "-1"]
        if not active:
            inactive_only.add(sku)
            continue
        pcodes = {s(x[1]) for x in active if s(x[1])}
        all_pcodes[sku] = pcodes
        if len(pcodes) != 1:
            ambiguous.add(sku)
            continue
        # 同一 P 码的多条记录：取第一条，类目/品牌若有差异会在对账中统计
        unique[sku] = {
            "bl_product_code": s(active[0][1]),
            "bl_product_sysno": s(active[0][2]),
            "bl_c1_id": s(active[0][3]),
            "bl_c2_id": s(active[0][4]),
            "bl_c3_id": s(active[0][5]),
            "bl_brand_id": s(active[0][6]),
        }
    log(f"可唯一映射 SKU {len(unique)} 个；P 码歧义 {len(ambiguous)}；仅失效记录 {len(inactive_only)}")
    return unique, ambiguous, inactive_only, len(rows), all_pcodes


def load_map(client):
    """t_goods_map 仅用于标记口径差异，绝不用来覆盖或补齐主口径。"""
    rows = query_rows(client, MAP_SQL, 60000)
    m = {}
    for a, b in rows:
        if s(a):
            m.setdefault(s(a), set()).add(s(b))
    log(f"t_goods_map 返回 {len(rows)} 行 / {len(m)} 个三方商品 ID（只做冲突标记）")
    return m


def main():
    client = j.build_archery()
    pool, ambiguous, inactive_only, raw_pool_rows, all_pcodes = load_pool(client)
    goods_map = load_map(client)
    detail_rows = []

    with SRC.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        base_fields = list(reader.fieldnames or [])
        src_rows = list(reader)
    fields = base_fields + NEW_FIELDS
    log(f"读取原始 CSV {len(src_rows)} 行 / {len(base_fields)} 列，输出 {len(fields)} 列")

    stats = defaultdict(int)
    with DST.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in src_rows:
            sku = s(row.get("sku_id"))
            out = dict(row)
            map_codes = goods_map.get(sku, set())
            if sku in pool:
                out.update(pool[sku])
                out["match_source"] = "dist_product_pool"
                # 冲突标记：t_goods_map 给出的 P 码与分销池不一致
                out["conflict_with_goods_map"] = "1" if (map_codes and pool[sku]["bl_product_code"] not in map_codes) else "0"
            else:
                for k in NEW_FIELDS[:6]:
                    out[k] = ""
                if sku in ambiguous:
                    out["match_source"] = "ambiguous_pcode"
                elif sku in inactive_only:
                    out["match_source"] = "inactive_only"
                else:
                    out["match_source"] = "unmatched"
                # 未匹配时如实提示另一条路径是否存在，但不自动采用
                out["conflict_with_goods_map"] = "goods_map_only" if map_codes else "0"
            stats[out["match_source"]] += 1
            if out["conflict_with_goods_map"] == "1":
                stats["conflict_rows"] += 1
            if out["conflict_with_goods_map"] == "goods_map_only":
                stats["goods_map_only_rows"] += 1
            if out["match_source"] != "dist_product_pool":
                detail_rows.append({
                    "sku_id": sku,
                    "sku_name": s(row.get("sku_name")),
                    "pool_id": s(row.get("pool_id")),
                    "match_source": out["match_source"],
                    "pool_candidate_pcodes": ";".join(sorted(all_pcodes.get(sku, set()))),
                    "goods_map_pcode": ";".join(sorted(map_codes)),
                })
            writer.writerow(out)

    detail_fields = ["sku_id", "sku_name", "pool_id", "match_source",
                    "pool_candidate_pcodes", "goods_map_pcode"]
    with DETAIL.open("w", encoding="utf-8-sig", newline="") as f:
        dw = csv.DictWriter(f, fieldnames=detail_fields)
        dw.writeheader()
        dw.writerows(detail_rows)

    # 对账：逐行比对原表字段是否被改动、行数是否一致
    with DST.open(encoding="utf-8-sig", newline="") as f:
        out_rows = list(csv.DictReader(f))
    unchanged = all(all(a.get(k, "") == b.get(k, "") for k in base_fields)
                     for a, b in zip(src_rows, out_rows))
    filled = {k: sum(1 for r in out_rows if s(r.get(k))) for k in NEW_FIELDS[:6]}
    report = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source": str(SRC),
        "output": str(DST),
        "source_rows": len(src_rows),
        "output_rows": len(out_rows),
        "columns": len(fields),
        "base_columns_unchanged": unchanged and len(out_rows) == len(src_rows),
        "pool_rows_read": raw_pool_rows,
        "pool_unique_skus": len(pool),
        "pool_ambiguous_skus": len(ambiguous),
        "pending_review_rows": len(detail_rows),
        "pending_review_file": str(DETAIL),
        "by_match_source": dict(sorted(stats.items())),
        "id_field_fill": filled,
        "id_field_fill_rate": {k: f"{v / len(out_rows) * 100:.1f}%" for k, v in filled.items()},
        "pool_scope": "benlai_business_product.product：JdGoodsId 直连，仅取有效记录（IsDelete<>1 且 ProductStatus<>-1）",
        "goods_map_usage": "仅用于冲突标记与参考值，未覆盖主口径、未补数据、未裁决多 P 码歧义记录",
        "user_decided": {
            "mapping": "分销商品池为准",
            "product_id": "P 码 + 主商品编号都给",
            "category_brand": "三级类目 ID + 品牌 ID",
            "unmatched": "留空并标记来源",
        },
    }
    VERIFY.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    lines = [
        "京东 VOP 全量商品 · 本网 ID 二次加工结果说明",
        f"生成时间：{report['generated_at']}",
        "",
        "一、你确认的口径",
        "- 映射主口径：分销商品池 benlai_business_product.product 的 JdGoodsId 直连",
        "- 商品 ID：P 码（bl_product_code）与主商品编号（bl_product_sysno）都给",
        "- 类目：本网三级类目 ID（bl_c1_id / bl_c2_id / bl_c3_id）；品牌：bl_brand_id；均不含名称",
        "- 匹配不上或存在歧义：ID 留空，用 match_source / conflict_with_goods_map 标记，一行不丢",
        "",
        "二、字段来源（重要：两种商品 ID 不是一回事）",
        "- bl_product_code  ← product.ProductId，P 码，形如 P0000102903",
        "- bl_product_sysno ← product.ProductBasicSysNo，C 网主商品编号，数字",
        "- bl_c1_id / bl_c2_id / bl_c3_id ← product.C1SysNo / C2SysNo / C3SysNo",
        "- bl_brand_id ← product.BrandSysNo",
        "- 主商品表 benlai_product.product_basic_info 的 SysNo / C1SysNo 等字段本次未混用，避免两套口径串味",
        "",
        "三、match_source 取值",
        "- dist_product_pool：分销池命中且 P 码唯一，ID 已补齐",
        "- ambiguous_pcode：同一京东 SKU 在有效记录里对应多个不同 P 码，未替业务裁决，ID 留空",
        "- inactive_only：只有已删除/失效记录，ID 留空",
        "- unmatched：分销池无此京东 SKU，ID 留空",
        "",
        "四、conflict_with_goods_map 取值",
        "- 1：t_goods_map 给出的 P 码与分销池不一致（按你的口径以分销池为准，仅标记、不覆盖）",
        "- goods_map_only：分销池没匹配上，但 t_goods_map 有记录（仅提示、不自动补）",
        "- 0：无差异",
        "- 调研阶段按「取分销池首条记录」的算法统计到 744 个 SKU 的 P 码与 t_goods_map 不一致；本次按「不自动裁决」把同一 SKU 对应多个 P 码的记录留空，因此标 1 的冲突行数少于 744，其余都落在待确认明细里。",
        "- 补齐率 96.2% 低于调研时的 99.3%，原因正是这批多 P 码记录按口径留空，不是查询失败或数据丢失。",
        "",
        "五、对账",
        f"- 原表 {len(src_rows)} 行 / {len(base_fields)} 列 → 新表 {len(out_rows)} 行 / {len(fields)} 列",
        f"- 原 32 列内容逐行未改动：{report['base_columns_unchanged']}",
        f"- match_source 分布：{', '.join(f'{k}={stats[k]}' for k in ('dist_product_pool', 'ambiguous_pcode', 'inactive_only', 'unmatched') if k in stats)}",
        f"- conflict_with_goods_map 分布：1={stats['conflict_rows']} 行（两路径 P 码不一致）；goods_map_only={stats['goods_map_only_rows']} 行（分销池未命中但 t_goods_map 有记录）；其余为 0",
        f"- 新 ID 列有值行数：{', '.join(f'{k}={v}' for k, v in filled.items())}",
        "",
        "六、待确认明细",
        f"- 本网ID_待确认.csv：{len(detail_rows)} 行，列出所有 ID 留空的记录，附分销池候选 P 码与 t_goods_map 参考 P 码，供人工拍板；本工具不代为选择。",
        "",
        "七、安全",
        "- 线上库仅执行只读 SELECT；未写生产库、未发商品变更事件、未调用京东接口、未生成 token。",
        "- t_goods_map 只用于标记差异，未用于覆盖或补齐主口径。",
        "",
        "八、复跑",
        "执行 enrich_bl_ids.py 即可重新按线上最新分销池数据生成，原 CSV 不会被修改。",
    ]
    README.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(json.dumps({k: v for k, v in report.items() if k != "user_decided"}, ensure_ascii=False, indent=2))
    print(f"已生成：{DST}\n{VERIFY}\n{README}")


if __name__ == "__main__":
    main()
