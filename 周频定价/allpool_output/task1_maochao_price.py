#!/usr/bin/env python3
"""任务1：对本网↔猫超已对齐的品，用猫超券后价折算出本网品的市场价。

规则（用户拍板）：
- 猫超价取券后价（coupon_price，SPU 代表价）；券后价无值回退一口价
- 不同规格折算：单件价 = 猫超价 ÷ 猫超件数；本网预估市场价 = 单件价 × 本网件数
  （双方件数用 parse_pack 从名称解析；同规格时即原价 passthrough）
- 产出：allpool_output/任务1_猫超市场价.jsonl
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "pricing_output"))
from build_similarity import parse_pack

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
from latest import latest
ALIGN = HERE / "对齐_本网猫超.jsonl"
PRICE = latest("supply_pull/output/猫超价格_全量_*.jsonl", "ALLPOOL_MC_PRICE")
DETAIL = latest("supply_pull/output/猫超详情_全量_*.jsonl", "ALLPOOL_MC_DETAIL")
INFO = latest("pricing_output/源数据_标准池商品信息_*.jsonl", "ALLPOOL_INFO")
OUT = HERE / "任务1_猫超市场价.jsonl"


def yuan(v):
    try:
        return round(float(v) / 100, 2) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def main():
    price = {}
    for l in PRICE.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            price[o["product_id"]] = o
    detail = {}
    from latest import latest, merged_lines
    detail = {}
    for l in merged_lines("supply_pull/output/猫超详情_全量_*.jsonl", "ALLPOOL_MC_DETAIL"):
        o = json.loads(l)
        detail[o["product_id"]] = o["data"]
    info = {json.loads(l)["productSysNo"]: json.loads(l) for l in INFO.open(encoding="utf-8") if l.strip()}
    # pbasic → 商品名（标准池 jsonl 带 productSysNo）
    pbasic2name = {}
    pool = latest("pricing_output/标准池_四站全上架有库存_*.jsonl", "ALLPOOL_POOL")
    for l in pool.open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            inf = info.get(o["productSysNo"], {})
            pbasic2name[o["productBasicSysNo"]] = inf.get("name") or ""

    n = ok = 0
    with OUT.open("w", encoding="utf-8") as f:
        for l in ALIGN.open(encoding="utf-8"):
            if not l.strip():
                continue
            o = json.loads(l)
            n += 1
            pid = o["maochao_pid"]
            pr = price.get(pid) or {}
            d = detail.get(pid) or {}
            # 券后价优先，回退一口价
            mc_price = pr.get("coupon_price")
            src = "猫超券后价"
            if mc_price is None:
                mc_price = yuan(d.get("reserve_price"))
                src = "猫超一口价(券后无值)"
            if mc_price is None:
                continue
            mc_name = d.get("product_title") or ""
            bl_name = pbasic2name.get(o["pbasic"]) or ""
            mc_pack = parse_pack(mc_name)
            bl_pack = parse_pack(bl_name)
            mc_cnt = mc_pack.get("count") or 1
            bl_cnt = bl_pack.get("count") or 1
            if mc_pack.get("value") and bl_pack.get("value") and mc_pack.get("base") == bl_pack.get("base"):
                # 基础规格相同：单件价 × 本方件数
                est = round(mc_price / mc_cnt * bl_cnt, 2)
            elif mc_pack.get("value") and bl_pack.get("value"):
                # 基础规格不同：按单位净含量折算
                mc_total = mc_pack["value"] * mc_cnt
                bl_total = bl_pack["value"] * bl_cnt
                est = round(mc_price / mc_total * bl_total, 2) if mc_total else None
            else:
                est = round(mc_price / mc_cnt * bl_cnt, 2)
            if est is None:
                continue
            ok += 1
            f.write(json.dumps({
                "pbasic": o["pbasic"], "maochao_pid": pid,
                "via": o.get("via"), "score": o.get("score"),
                "bl_name": bl_name, "mc_name": mc_name,
                "mc_price": mc_price, "price_src": src,
                "mc_spec": mc_pack.get("raw") or "", "bl_spec": bl_pack.get("raw") or "",
                "est_market_price": est,
            }, ensure_ascii=False) + "\n")
    print(json.dumps({"对齐品数": n, "成功折算": ok, "输出": OUT.name}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
