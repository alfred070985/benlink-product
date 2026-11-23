#!/usr/bin/env python3
import json
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CURRENT = ROOT / "online_product_costs.jsonl"
MAPPINGS = ROOT / "product_channel_mappings.jsonl"
PRICES = ROOT / "channel_product_price_records.jsonl"
PROGRESS = ROOT / "channel_price_fetch_progress.json"
OUTPUT = ROOT / "product_channel_median_costs.jsonl"
SUMMARY = ROOT / "product_channel_median_costs_summary.json"


def median(values):
    ordered = sorted(values)
    count = len(ordered)
    middle = count // 2
    if count % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / Decimal(2)


def main():
    progress = json.loads(PROGRESS.read_text(encoding="utf-8"))
    assert progress["coverage_complete"] is True
    pool = {}
    with CURRENT.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                pool[item["productId"]] = item
    prices = {}
    with PRICES.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                prices[int(item["ProductSysNo"])] = item
    tiers = {pid: {"分销渠道": [], "C网": [], "非集采全渠道": []} for pid in pool}
    mapped_basics = set()
    mapping_stats = Counter()
    with MAPPINGS.open(encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            item = json.loads(line)
            if item["IsDelete"]:
                continue
            pid = item["pool_product_id"]
            basic = int(item["ProductBasicSysNo"])
            mapped_basics.add(basic)
            product_sysno = int(item["ProductSysNo"])
            price = prices.get(product_sysno)
            if price is None:
                raise RuntimeError(f"Missing collected price: {product_sysno}")
            row = {
                "ProductSysNo": product_sysno,
                "WebSiteSysNo": int(item["WebSiteSysNo"]),
                "ChannelSysNo": int(item["ChannelSysNo"]),
                "SaleChannelSysNo": int(item["SaleChannelSysNo"]),
                "Status": int(item["Status"]),
                "UnitCost": Decimal(str(price["UnitCost"])),
            }
            mapping_stats["有效渠道商品记录"] += 1
            if row["ChannelSysNo"] == 5 and row["SaleChannelSysNo"] == 5:
                tiers[pid]["分销渠道"].append(row)
                mapping_stats["分销记录"] += 1
            if row["SaleChannelSysNo"] == 1:
                tiers[pid]["C网"].append(row)
                mapping_stats["C网记录"] += 1
            if row["SaleChannelSysNo"] != 2:
                tiers[pid]["非集采全渠道"].append(row)
                mapping_stats["非集采记录"] += 1

    stats = Counter()
    results = {}
    for pid in pool:
        selected_source = None
        selected = []
        for source in ("分销渠道", "C网", "非集采全渠道"):
            if tiers[pid][source]:
                selected_source = source
                selected = tiers[pid][source]
                break
        if not selected:
            stats["无有效渠道记录"] += 1
            result = {
                "product_id": pid,
                "productBasicSysNo": pool[pid]["productBasicSysNo"],
                "selected_source": "无有效渠道记录",
                "UnitCost": None,
                "selected_record_count": 0,
                "selected_website_count": 0,
                "selected_product_sysnos": "",
                "selected_websites": "",
                "unit_cost_min": None,
                "unit_cost_max": None,
            }
        else:
            costs = [row["UnitCost"] for row in selected]
            value = median(costs)
            sysnos = sorted(row["ProductSysNo"] for row in selected)
            websites = sorted(set(row["WebSiteSysNo"] for row in selected))
            result = {
                "product_id": pid,
                "productBasicSysNo": pool[pid]["productBasicSysNo"],
                "selected_source": selected_source,
                "UnitCost": format(value, "f"),
                "selected_record_count": len(selected),
                "selected_website_count": len(websites),
                "selected_product_sysnos": ",".join(map(str, sysnos)),
                "selected_websites": ",".join(map(str, websites)),
                "unit_cost_min": format(min(costs), "f"),
                "unit_cost_max": format(max(costs), "f"),
            }
            stats[selected_source] += 1
            stats["多记录取中位数" if len(selected) > 1 else "单记录"] += 1
            if value == 0:
                stats["中位数为0"] += 1
            elif value > 0:
                stats["中位数为正"] += 1
            else:
                stats["中位数为负"] += 1
        results[pid] = result
    with OUTPUT.open("w", encoding="utf-8") as f:
        for pid in sorted(results):
            f.write(json.dumps(results[pid], ensure_ascii=False) + "\n")
    summary = {
        "pool_products": len(pool),
        "mapped_products": len(mapped_basics),
        "products_without_mapping": len(pool) - len(mapped_basics),
        "mapping_stats": dict(mapping_stats),
        "selection_stats": dict(stats),
        "median_rule": "按UnitCost升序，奇数条取中间值，偶数条取中间两条算术平均；0值参与计算",
        "priority": ["分销渠道(ChannelSysNo=5,SaleChannelSysNo=5)", "C网(SaleChannelSysNo=1)", "全部排除集采(SaleChannelSysNo<>2)"],
        "active_rule": "IsDelete=false；不以Status上下架状态过滤",
        "time_filter": None,
        "output": str(OUTPUT),
    }
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
