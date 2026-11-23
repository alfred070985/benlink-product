#!/usr/bin/env python3
"""全货盘聚合表：三底表并集 + 同品组 + 统一直发档定价 + 空间率≥10%过滤。

骨架（每货源一行，用户拍板不合并）：
  本网 75,067：成本=ERP采用成本；市面价格=天镜/京东对标价(成品)，无则任务1猫超折算价
  京东VOP：有协议价+C端价的 32,878；成本=sale_price；市面价格=jd_price
  猫超：813,572 SKU 行；成本=供货价；市面价格=券后价(回退一口价)
统一直发档 2/5/8（用户拍板）；建议档=默认中档②，空间<10%降①，≥30%且档③≥10%升③（沿用拍板规则）。
group_id：并查集（本网↔猫超对齐 1,637 + 京东↔本网 16,429 + 京东↔猫超 117，传递成团）。
输出：allpool_output/全货盘聚合表_<日期>.csv
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
    """返回 (建议档序号0/1/2, 说明)。规则沿用拍板：默认中档；<10%降①；≥30%且档③≥10%升③。"""
    if not market or t[1] is None:
        return 1, "无市场价·默认中档"
    s2 = (market - t[1]) / market
    if s2 < FLOOR:
        return 0, f"中档空间{s2:.0%}<10%·降档①"
    s3 = (market - t[2]) / market if t[2] is not None else None
    if s2 >= HI and s3 is not None and s3 >= FLOOR:
        return 2, f"中档空间{s2:.0%}≥30%·升档③"
    return 1, f"中档空间{s2:.0%}·中档"


def main():
    # ---- 1. 本网行（读货盘一成品，完整沿用内部逻辑：直发2/5/8、仓发5/12/24 分开发）----
    print("读本网成品…", flush=True)
    bl_rows = []
    wb = openpyxl.load_workbook(ROOT / "pricing_output" / "分销标准池定价_2026-11-23.xlsx",
                                read_only=True, data_only=True)
    ws = wb["01_定价结果"]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    ix = {k: hdr.index(k) for k in ("商品编号", "商品名称", "二级类目", "发货类型", "采用成本(元)",
                                    "档①输出价", "档②输出价", "档③输出价",
                                    "建议档", "建议输出价(元)", "对标价(元)", "对标来源")}
    task1 = {}
    for l in (HERE / "任务1_猫超市场价.jsonl").open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            task1[str(o["pbasic"])] = o
    # 商品编号(productSysNo) → productBasicSysNo 映射（任务1/对齐的key是pbasic）
    psys2pbasic = {}
    for l in (ROOT / "pricing_output" / "标准池_四站全上架有库存_2026-11-19.jsonl").open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            psys2pbasic[str(o["productSysNo"])] = str(o["productBasicSysNo"])
    for r in it:
        if not r or not r[0]:
            continue
        pid = str(r[ix["商品编号"]]).strip()
        market = r[ix["对标价(元)"]]
        msrc = r[ix["对标来源"]] or ""
        pbasic = psys2pbasic.get(pid)
        if (market is None or "ERP基准价" in str(msrc)) and pbasic in task1:
            t1 = task1[pbasic]
            market = t1["est_market_price"]
            msrc = f"猫超折算({t1['price_src']})"
        bl_rows.append({
            "key": f"bl:{pbasic or pid}", "货源": "本网", "商品名": r[ix["商品名称"]],
            "品牌": "", "类目": r[ix["二级类目"]] or "", "规格": "",
            "条码": "", "成本": r[ix["采用成本(元)"]], "市面价格": market, "价格来源": msrc,
            "发货类型": r[ix["发货类型"]],
            "档①": r[ix["档①输出价"]], "档②": r[ix["档②输出价"]], "档③": r[ix["档③输出价"]],
            "建议档": r[ix["建议档"]], "建议输出价": r[ix["建议输出价(元)"]],
        })
    wb.close()
    print(f"本网 {len(bl_rows)} 行（任务1补充 {sum(1 for b in bl_rows if '猫超折算' in str(b['价格来源']))}）", flush=True)

    # ---- 2. 京东行 ----
    print("读京东VOP…", flush=True)
    jd_rows = []
    with (ROOT / "jdvop_pull" / "output" / "本来生活VOP_goods_本网ID.csv").open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            cost = r.get("sale_price")
            market = r.get("jd_price")
            try:
                cost = float(cost) if cost else None
            except ValueError:
                cost = None
            try:
                market = float(market) if market else None
            except ValueError:
                market = None
            if not cost or not market:
                continue
            jd_rows.append({
                "key": f"jd:{r['sku_id']}", "货源": "京东VOP", "商品名": r.get("sku_name"),
                "品牌": r.get("brand_name") or "", "类目": r.get("category") or "", "规格": "",
                "条码": "", "成本": cost, "市面价格": market, "价格来源": "京东C端价",
            })
    print(f"京东 {len(jd_rows)} 行", flush=True)

    # ---- 3. 猫超行（SKU级宽表）----
    print("读猫超宽表…", flush=True)
    mc_rows = []
    with (ROOT / "jdvop_pull" / "output" / "猫超商品宽表SKU级_2026-11-21.csv").open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            cost = r.get("channel_price_sku") or r.get("channel_price_spu")
            market = r.get("coupon_price_sku") or r.get("coupon_price_spu")
            if not market:
                market = r.get("reserve_price_yuan")
            try:
                cost = float(cost) if cost else None
            except ValueError:
                cost = None
            try:
                market = float(market) if market else None
            except ValueError:
                market = None
            if not cost:
                continue
            mc_rows.append({
                "key": f"mc:{r['product_id']}:{r.get('sku_id') or ''}", "货源": "猫超",
                "商品名": r.get("product_title"), "品牌": r.get("brand") or "",
                "类目": r.get("category") or "", "规格": r.get("sku_properties") or "",
                "条码": r.get("sku_barcode") or "", "成本": cost, "市面价格": market,
                "价格来源": "猫超券后价" if r.get("coupon_price_sku") or r.get("coupon_price_spu") else "猫超一口价",
            })
    print(f"猫超 {len(mc_rows)} 行", flush=True)

    # ---- 4. 并查集同品组 ----
    dsu = DSU()
    for src in ("对齐_本网猫超.jsonl", "对齐_京东_bl.jsonl", "对齐_京东_mc.jsonl"):
        for l in (HERE / src).open(encoding="utf-8"):
            if not l.strip():
                continue
            o = json.loads(l)
            if src == "对齐_本网猫超.jsonl":
                dsu.union(f"bl:{o['pbasic']}", f"mc:{o['maochao_pid']}:")
            elif src == "对齐_京东_bl.jsonl":
                dsu.union(f"jd:{o['jd_sku']}", f"bl:{o['target_id']}")
            else:
                dsu.union(f"jd:{o['jd_sku']}", f"mc:{o['target_id']}:")
    # key 前缀归一（mc key 带 sku_id 后缀，对齐按 pid 前缀连）
    print("并查集就绪", flush=True)

    # ---- 5. 组装+统一直发档+过滤 ----
    fields = ["group_id", "货源", "商品名", "品牌", "类目", "规格", "条码",
              "发货类型", "加价档类型", "成本", "档①", "档②", "档③",
              "建议档", "建议输出价", "市面价格", "价格来源", "空间率", "建议说明"]
    n_total = n_keep = 0
    by_src = {}
    with OUT.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for row in bl_rows + jd_rows + mc_rows:
            n_total += 1
            if row["货源"] == "本网":
                # 本网：完整沿用货盘一（直发2/5/8、仓发5/12/24 + 内部建议档）
                t = (row["档①"], row["档②"], row["档③"])
                out_price = row["建议输出价"]
                note = "沿用货盘一建议档"
                tier_type = "直发2/5/8" if row.get("发货类型") == "直发" else "仓发5/12/24"
                ship = row.get("发货类型") or ""
            else:
                # 外部货源（京东VOP/猫超）：对方发货=直发档 2/5/8
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
            parts = row["key"].split(":")
            gid = dsu.find(":".join(parts[:2]))
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
                      "输出": OUT.name}, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
