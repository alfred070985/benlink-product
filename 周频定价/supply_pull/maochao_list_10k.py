#!/usr/bin/env python3
"""猫超产品列表批量拉取（默认1万个）+ 标题关键词品类分布统计。

list 接口每页上限 100；品类分布按标题关键词粗分（详情接口字段更多，按需再回查）。
输出：output/猫超产品_list_<n>个_<时间戳>.jsonl / .csv + 品类分布打印
"""
import csv
import json
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from maochao_pull import load_creds, top_call, extract_items, flatten, CHANNEL, SUB_CHANNEL, OUT_DIR

# 标题关键词 → 品类（粗分，供分布参考）
CAT_RULES = [
    ("图书", ["书", "ISBN", "出版社", "著", "小说", "绘本", "读物", "漫画", "教材", "丛书"]),
    ("宠物", ["猫粮", "狗粮", "宠物", "猫条", "猫砂"]),
    ("酒水", ["酒", "啤", "葡萄", "茅台", "五粮液", "洋酒"]),
    ("美妆个护", ["洗发", "沐浴", "洗面奶", "面霜", "精华", "口红", "面膜", "牙膏", "卫生巾", "洗衣", "护肤", "香水", "卸妆"]),
    ("母婴", ["奶粉", "尿裤", "奶瓶", "婴儿", "宝宝", "辅食", "孕妈", "纸尿裤", "拉拉裤"]),
    ("食品生鲜", ["饼", "坚果", "茶", "咖啡", "油", "大米", "面粉", "挂面", "奶", "饮料", "水果", "肉", "菜", "盐", "糖", "酱", "零食", "糕", "蜜", "粮", "蛋", "鱼", "虾", "菇", "酸菜", "肠", "卤"]),
    ("家居清洁", ["纸巾", "抽纸", "卷纸", "垃圾袋", "洗洁精", "清洁", "拖把", "收纳", "保鲜", "碗", "杯", "锅"]),
    ("家电数码", ["手机", "耳机", "电池", "充电", "电饭煲", "键盘", "鼠标", "插排", "风扇", "吹风", "剃须", "电动牙刷"]),
    ("服饰家纺", ["T恤", "外套", "裤", "鞋", "袜", "内衣", "被", "床单", "枕", "毛巾", "绒"]),
    ("玩具文具", ["玩具", "乐高", "积木", "笔", "文具", "本册", "拼图", "模型"]),
    ("保健", ["维生素", "蛋白", "保健", "钙", "益生菌", "酵素"]),
]


def cat_of(title):
    for cat, words in CAT_RULES:
        if any(w in title for w in words):
            return cat
    return "其他"


def main():
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 10000
    page_size = 100
    creds = load_creds()
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    raw_path = OUT_DIR / f"猫超产品_list_{limit}个_{stamp}.jsonl"
    csv_path = OUT_DIR / f"猫超产品_list_{limit}个_{stamp}.csv"

    rows, page = [], 1
    cats = Counter()
    scroll_id = None
    t0 = time.time()
    with raw_path.open("w", encoding="utf-8") as raw:
        while len(rows) < limit:
            # 分页深度限制在 5000（第51页起服务端报错），scroll 全量模式无此限制
            biz = {"channel_code": CHANNEL, "sub_channel_code": SUB_CHANNEL,
                   "page_size": page_size, "is_scroll": True}
            if scroll_id:
                biz["scroll_id"] = scroll_id
            else:
                biz["page_no"] = 1
            resp = None
            for attempt in range(8):
                try:
                    resp = top_call(creds, biz)
                    if "error_response" in resp:
                        time.sleep(min(2 * (attempt + 1), 15))
                        continue
                    break
                except Exception:
                    if attempt == 7:
                        raise
                    time.sleep(min(2 * (attempt + 1), 15))
            if resp is None or "error_response" in resp:
                print("接口报错(已重试8次):", json.dumps((resp or {}).get("error_response", {}), ensure_ascii=False)[:300], flush=True)
                break
            r = resp.get("alibaba_ascp_channel_distributor_product_list_response", {})
            data = (r.get("product_list_response") or {}).get("data") or {}
            items, total = extract_items(resp)
            scroll_id = data.get("scroll_id") or scroll_id
            if not items:
                break
            for p in items:
                raw.write(json.dumps(p, ensure_ascii=False) + "\n")
                f = flatten(p)
                f["category_guess"] = cat_of(f.get("product_title") or "")
                cats[f["category_guess"]] += 1
                rows.append(f)
            if page % 20 == 0 or len(items) < page_size:
                el = time.time() - t0
                print(f"page {page}: 累计 {len(rows)}/{limit} (total={total}) {el:.0f}s", flush=True)
            if len(items) < page_size:
                break
            page += 1
            time.sleep(0.4)

    rows = rows[:limit]
    fields = ["product_id", "product_title", "sales_modes", "picture", "gmt_modified", "category_guess"]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)

    print("\n=== 品类分布（标题关键词粗分）===")
    for cat, n in cats.most_common():
        print(f"  {cat}: {n} ({n/len(rows):.1%})")
    print(json.dumps({"拉取数": len(rows), "csv": csv_path.name, "raw": raw_path.name,
                      "耗时秒": round(time.time() - t0, 1)}, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
