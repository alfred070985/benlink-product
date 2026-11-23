#!/usr/bin/env python3
"""合并猫超详情 + 价格 → SKU 级全量宽表（每个 SKU 一行，四价齐全）。

用户拍板：多规格商品必须按 SKU 展开（不同 SKU 不同价格），行数 > 14.7 万。
SKU 编号在详情/价格接口间完全对齐（已验证），按 sku_id 精确匹配。

每行 = SPU 信息（重复横铺） + SKU 信息 + SKU 四价：
  一口价(SKU) / 渠道供货价(SKU) / 券后价(SKU) / 建议零售价(猫超未启用,空列)
  SPU 级供货价/券后价同时带上（SKU 无价时兜底看 SPU 列）。
无 SKU 明细的商品退化为 1 行（SPU 级价格）。

输入（取最新）：猫超详情_全量_*.jsonl / 猫超价格_全量_*.jsonl
输出：猫超商品宽表SKU级_<日期>.csv
"""
import csv
import glob
import json
from collections import Counter
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
TODAY = datetime.now().strftime("%Y-%m-%d")
OUT_CSV = OUT_DIR / f"猫超商品宽表SKU级_{TODAY}.csv"


def str_list(v):
    if isinstance(v, dict):
        return v.get("string") or []
    return v or []


def latest(pat):
    fs = sorted(glob.glob(str(OUT_DIR / pat)))
    return Path(fs[-1]) if fs else None


def yuan(v):
    try:
        return round(float(v) / 100, 2) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def load_details():
    """周频增量：以最新 list 为有效品范围（自动剔除已下架），详情取所有历史文件并集（新品覆盖旧数据）。
    找不到当周 list 时退化为不过滤（详情并集全量）。"""
    list_path = latest("猫超产品_list全量*.jsonl")
    valid = None
    if list_path:
        valid = set()
        for l in list_path.open(encoding="utf-8"):
            if l.strip():
                valid.add(json.loads(l)["product_id"])
    out = {}
    srcs = sorted(glob.glob(str(OUT_DIR / "猫超详情_全量_*.jsonl")))
    for fp in srcs:
        for l in open(fp, encoding="utf-8"):
            if l.strip():
                o = json.loads(l)
                if valid is None or o["product_id"] in valid:
                    out[o["product_id"]] = o["data"]
    scope = f"+{list_path.name}" if list_path else "(无list,不过滤)"
    return f"{len(srcs)}个详情文件{scope}", out


def load_prices():
    path = latest("猫超价格_全量_*.jsonl")
    out = {}
    if not path:
        return "(无)", out
    for l in path.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            out[o["product_id"]] = o
    return path.name, out


def main():
    dname, details = load_details()
    pname, prices = load_prices()
    print(f"详情 {len(details)} 品（{dname}） | 价格 {len(prices)} 品（{pname}）", flush=True)

    fields = [
        "product_id", "product_title", "product_sub_title",
        "brand", "brand_id", "category", "category_id",
        "sku_id", "sku_properties", "sku_barcode", "reserve_price_yuan",
        "channel_price_sku", "coupon_price_sku", "advise_sale_price",
        "channel_price_spu", "coupon_price_spu",
        "供货价折扣率",  # SKU 供货价 ÷ SKU 一口价；无 SKU 价时用 SPU 比 SPU
        "券后毛利比",    # SKU 券后价 ÷ SKU 供货价；>1 有毛利空隙，<1 倒挂
        "tax_rate", "tax_code", "sale_unit",
        "temperature_requirement", "sales_modes", "supplier_user_id",
        "sku_count", "picture", "sku_picture",
        "channel_code", "sub_channel_code",
    ]

    stat = Counter()
    n_rows = 0
    with OUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for pid in sorted(details):
            d = details[pid]
            pr = prices.get(pid) or {}
            pics = str_list(d.get("pictures"))
            mats = {}
            try:
                mats = json.loads(d.get("materials_map") or "{}")
            except Exception:
                pass
            psd = d.get("product_sku_detail") or {}
            skus = psd.get("product_sku_detail_d_t_o") or [] if isinstance(psd, dict) else []
            ch_sku = pr.get("channel_price_sku") or {}
            cp_sku = pr.get("coupon_price_sku") or {}
            ch_spu = pr.get("channel_price")
            cp_spu = pr.get("coupon_price")
            base = {
                "product_id": pid,
                "product_title": d.get("product_title"),
                "product_sub_title": d.get("product_sub_title"),
                "brand": d.get("brand"),
                "brand_id": d.get("brand_id"),
                "category": d.get("category"),
                "category_id": d.get("category_id"),
                "tax_rate": d.get("tax_rate"),
                "tax_code": d.get("tax_code"),
                "sale_unit": mats.get("sale_unit"),
                "temperature_requirement": d.get("temperature_requirement"),
                "sales_modes": ",".join(str_list(d.get("sales_modes"))),
                "supplier_user_id": d.get("supplier_user_id"),
                "sku_count": len(skus),
                "picture": pics[0] if pics else None,
                "channel_code": d.get("channel_code"),
                "sub_channel_code": d.get("sub_channel_code"),
                "channel_price_spu": ch_spu,
                "coupon_price_spu": cp_spu,
            }
            lines = []
            if skus:
                for s in skus:
                    sid = str(s.get("sku_id"))
                    rp = yuan(s.get("reserve_price"))
                    ch = ch_sku.get(sid)
                    cp = cp_sku.get(sid)
                    lines.append({**base,
                        "sku_id": sid,
                        "sku_properties": " | ".join(str_list(s.get("properties"))),
                        "sku_barcode": s.get("barcode"),
                        "reserve_price_yuan": rp,
                        "channel_price_sku": ch,
                        "coupon_price_sku": cp,
                        "advise_sale_price": None,
                        "sku_picture": (str_list(s.get("picture")) or [None])[0],
                        "供货价折扣率": round(ch / rp, 4) if (ch and rp) else
                                      (round(ch_spu / rp, 4) if (ch_spu and rp) else None),
                        "券后毛利比": round(cp / ch, 4) if (cp and ch) else
                                    (round(cp_spu / ch_spu, 4) if (cp_spu and ch_spu) else None),
                    })
            else:
                rp = yuan(d.get("reserve_price"))
                lines.append({**base,
                    "sku_id": None,
                    "sku_properties": None,
                    "sku_barcode": (str_list(d.get("barcode_list")) or [None])[0],
                    "reserve_price_yuan": rp,
                    "channel_price_sku": None,
                    "coupon_price_sku": None,
                    "advise_sale_price": None,
                    "sku_picture": None,
                    "供货价折扣率": round(ch_spu / rp, 4) if (ch_spu and rp) else None,
                    "券后毛利比": round(cp_spu / ch_spu, 4) if (cp_spu and ch_spu) else None,
                })
            for row in lines:
                w.writerow(row)
                n_rows += 1
                if row["reserve_price_yuan"]: stat["一口价"] += 1
                if row["channel_price_sku"] or (row["channel_price_spu"] and not skus): stat["供货价"] += 1
                if row["coupon_price_sku"] or (row["coupon_price_spu"] and not skus): stat["券后价"] += 1
                if row["供货价折扣率"]: stat["折扣率可算"] += 1

    print(json.dumps({"宽表总行数": n_rows, "SPU数": len(details), "价格覆盖行": dict(stat),
                      "输出": OUT_CSV.name}, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
