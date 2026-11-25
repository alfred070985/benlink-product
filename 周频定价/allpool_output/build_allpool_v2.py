#!/usr/bin/env python3
"""全货盘聚合表 v2：四源并集（本网纯池 + 京东VOP + 猫超 + 东方甄选）
+ 向量对齐 group_id + 统一直发档 + 空间率≥10%过滤。

相对 build_allpool.py 的变更：
  1. 本网行读纯池重建货盘一（58,423，ERP剔16,644京东品）——京东品不再双算
  2. 新增东方甄选第四货源（成本=purchasePrice, 市面价=retailPrice,
     统一直发档2/5/8, 只留canBuy=True）——用户2026-11-25拍板
  3. group_id 升级为 6 组向量对齐采信边（bl/jd/mc/eb 两两）
  4. 本网行对标价已含向量新价（货盘一v2产出），无需再读任务1折算

保持不变：
  - 每货源一行不合并；本网沿用内部档位(直发2/5/8、仓发5/12/24)
  - 外部货源对方发货=直发档2/5/8；选档默认中档②，<10%降①，≥30%且档③≥10%升③
  - 空间率<10%过滤
"""
import csv
import json
from datetime import datetime
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
TODAY = datetime.now().strftime("%Y-%m-%d")
OUT = HERE / f"全货盘聚合表_{TODAY}.csv"
RATIOS = (0.02, 0.05, 0.08)
HI, FLOOR = 0.30, 0.10


class DSU:
    def __init__(self):
        self.p = {}

    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[rb] = ra


def tiers(cost):
    if not cost:
        return [None, None, None]
    return [round(cost * (1 + r), 2) for r in RATIOS]


def suggest(t, market):
    if not market or t[1] is None:
        return 1, "无市场价·默认中档"
    s2 = (market - t[1]) / market
    if s2 < FLOOR:
        return 0, f"中档空间{s2:.0%}<10%·降档①"
    s3 = (market - t[2]) / market if t[2] is not None else None
    if s2 >= HI and s3 is not None and s3 >= FLOOR:
        return 2, f"中档空间{s2:.0%}≥30%·升档③"
    return 1, f"中档空间{s2:.0%}·中档"


def fnum(v):
    try:
        return float(v) if v not in (None, "") else None
    except (ValueError, TypeError):
        return None


def main():
    from latest import latest
    # 源路径全走 latest()（环境变量可覆盖）
    BL_XLSX = latest("pricing_output/分销标准池定价_*.xlsx", "ALLPOOL_BL_XLSX")
    BL_POOL = latest("pricing_output/标准池_四站全上架有库存_*.jsonl", "ALLPOOL_POOL")
    VOP_CSV = latest("supply_pull/output/本来生活VOP_goods_本网ID.csv", "ALLPOOL_VOP")
    MC_WIDE = latest("supply_pull/output/猫超商品宽表SKU级_*.csv", "ALLPOOL_MC_WIDE")
    EB_JSONL = latest("supply_pull/output/东方甄选商品_*.jsonl", "ALLPOOL_EB")
    print(f"货盘一成品: {BL_XLSX.name}", flush=True)

    # ---- 1. 本网行（纯池货盘一，完整沿用内部档位）----
    bl_rows = []
    wb = openpyxl.load_workbook(BL_XLSX, read_only=True, data_only=True)
    ws = wb["01_定价结果"]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    ix = {k: hdr.index(k) for k in ("商品编号", "商品名称", "二级类目", "发货类型", "采用成本(元)",
                                    "档①输出价", "档②输出价", "档③输出价",
                                    "建议档", "建议输出价(元)", "对标价(元)", "对标来源")}
    psys2pbasic = {}
    for l in BL_POOL.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            psys2pbasic[str(o["productSysNo"])] = str(o["productBasicSysNo"])
    for r in it:
        if not r or not r[0]:
            continue
        pid = str(r[ix["商品编号"]]).strip()
        pbasic = psys2pbasic.get(pid) or pid
        bl_rows.append({
            "key": f"bl:{pbasic}", "货源": "本网", "商品名": r[ix["商品名称"]],
            "品牌": "", "类目": r[ix["二级类目"]] or "", "规格": "", "条码": "",
            "成本": r[ix["采用成本(元)"]], "市面价格": r[ix["对标价(元)"]],
            "价格来源": r[ix["对标来源"]] or "", "发货类型": r[ix["发货类型"]],
            "档①": r[ix["档①输出价"]], "档②": r[ix["档②输出价"]], "档③": r[ix["档③输出价"]],
            "建议档": r[ix["建议档"]], "建议输出价": r[ix["建议输出价(元)"]],
        })
    wb.close()
    print(f"本网 {len(bl_rows)} 行", flush=True)

    # ---- 2. 京东行 ----
    jd_rows = []
    with VOP_CSV.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            cost, market = fnum(r.get("sale_price")), fnum(r.get("jd_price"))
            if not cost or not market:
                continue
            jd_rows.append({
                "key": f"jd:{r['sku_id']}", "货源": "京东VOP", "商品名": r.get("sku_name"),
                "品牌": r.get("brand_name") or "", "类目": r.get("category") or "",
                "规格": "", "条码": "", "成本": cost, "市面价格": market,
                "价格来源": "京东C端价",
            })
    print(f"京东 {len(jd_rows)} 行", flush=True)

    # ---- 3. 猫超行（SKU级宽表）----
    mc_rows = []
    with MC_WIDE.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            cost = r.get("channel_price_sku") or r.get("channel_price_spu")
            market = r.get("coupon_price_sku") or r.get("coupon_price_spu") or r.get("reserve_price_yuan")
            cost, market = fnum(cost), fnum(market)
            if not cost:
                continue
            mc_rows.append({
                "key": f"mc:{r['product_id']}:{r.get('sku_id') or ''}", "货源": "猫超",
                "商品名": r.get("product_title"), "品牌": r.get("brand") or "",
                "类目": r.get("category") or "", "规格": r.get("sku_properties") or "",
                "条码": r.get("sku_barcode") or "", "成本": cost, "市面价格": market,
                "价格来源": "猫超券后价" if (r.get("coupon_price_sku") or r.get("coupon_price_spu")) else "猫超一口价",
            })
    print(f"猫超 {len(mc_rows)} 行", flush=True)

    # ---- 4. 东方甄选行（新增第四源，canBuy=True）----
    eb_rows = []
    for l in EB_JSONL.open(encoding="utf-8"):
        if not l.strip():
            continue
        r = json.loads(l)
        if not r.get("canBuy"):
            continue
        cost, market = fnum(r.get("purchasePrice")), fnum(r.get("retailPrice"))
        if not cost or not market:
            continue
        eb_rows.append({
            "key": f"eb:{r['specOuterId']}", "货源": "东方甄选",
            "商品名": r.get("displayGoodsName") or r.get("goodsName") or "",
            "品牌": "", "类目": "", "规格": r.get("specSkuProperties") or "", "条码": "",
            "成本": cost, "市面价格": market, "价格来源": "东方甄选零售价",
        })
    print(f"东方甄选 {len(eb_rows)} 行", flush=True)

    # ---- 5. 并查集同品组（6组向量对齐，只取采信边）----
    dsu = DSU()
    # 本网侧只保留纯池 pbasic（被剔的16,644京东品不得当root污染group_id）
    pure_pb = {r["key"][3:] for r in bl_rows if r["key"].startswith("bl:")}
    # 对齐文件: (文件名, a的前缀, b的前缀)。b为尾侧时按各源ID归key
    align_specs = [
        ("对齐_向量_bl_jd.jsonl", "bl", "jd"),
        ("对齐_向量_bl_mc.jsonl", "bl", "mc"),
        ("对齐_向量_bl_eb.jsonl", "bl", "eb"),
        ("对齐_向量_jd_mc.jsonl", "jd", "mc"),
        ("对齐_向量_jd_eb.jsonl", "jd", "eb"),
        ("对齐_向量_mc_eb.jsonl", "mc", "eb"),
    ]
    edge_n = 0
    for fn, pa, pb in align_specs:
        fp = HERE / fn
        if not fp.exists():
            print(f"  跳过缺失 {fn}", flush=True)
            continue
        for l in fp.open(encoding="utf-8"):
            if not l.strip():
                continue
            o = json.loads(l)
            if o.get("tier") != "采信":
                continue
            # mc key 带 :sku_id 后缀，对齐里 b 是 product_id → 归一到 pid 前缀
            ka = f"{pa}:{o['a']}"
            kb = f"{pb}:{o['b']}"
            # 本网侧节点若不在纯池（即被剔京东品），跳过该边，避免脏root
            if pa == "bl" and o["a"] not in pure_pb:
                continue
            if pb == "bl" and o["b"] not in pure_pb:
                continue
            dsu.union(ka, kb)
            edge_n += 1
    print(f"并查集: {edge_n} 条采信边", flush=True)

    # ---- 6. 组装 + 档位 + 空间率过滤 ----
    fields = ["group_id", "货源", "商品名", "品牌", "类目", "规格", "条码",
              "发货类型", "加价档类型", "成本", "档①", "档②", "档③",
              "建议档", "建议输出价", "市面价格", "价格来源", "空间率", "建议说明"]
    n_total = n_keep = 0
    by_src = {}
    with OUT.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for row in bl_rows + jd_rows + mc_rows + eb_rows:
            n_total += 1
            if row["货源"] == "本网":
                t = (row["档①"], row["档②"], row["档③"])
                out_price = row["建议输出价"]
                note = "沿用货盘一建议档"
                tier_type = "直发2/5/8" if row.get("发货类型") == "直发" else "仓发5/12/24"
                ship = row.get("发货类型") or ""
            else:
                t = tiers(row["成本"])
                idx, note = suggest(t, row["市面价格"])
                out_price = t[idx]
                tier_type = "直发2/5/8"
                ship = "直发"
            space = (row["市面价格"] - out_price) / row["市面价格"] if (row["市面价格"] and out_price) else None
            if space is None or space < FLOOR:
                continue
            n_keep += 1
            by_src[row["货源"]] = by_src.get(row["货源"], 0) + 1
            # mc key 带 :sku_id 后缀，对齐按 product_id 建节点 → 只取前两段归一
            gid = dsu.find(":".join(row["key"].split(":")[:2]))
            w.writerow({
                "group_id": gid, "货源": row["货源"], "商品名": row["商品名"],
                "品牌": row["品牌"], "类目": row["类目"], "规格": row["规格"], "条码": row["条码"],
                "发货类型": ship, "加价档类型": tier_type,
                "成本": row["成本"], "档①": t[0], "档②": t[1], "档③": t[2],
                "建议档": row.get("建议档") or f"档{idx+1}", "建议输出价": out_price,
                "市面价格": row["市面价格"], "价格来源": row["价格来源"],
                "空间率": f"{space:.2%}", "建议说明": note,
            })
    print(json.dumps({"总行数": n_total, "过滤后": n_keep, "分货源": by_src,
                      "输出": str(OUT)}, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
