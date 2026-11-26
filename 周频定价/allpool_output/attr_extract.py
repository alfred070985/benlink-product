#!/usr/bin/env python3
"""商品属性结构化抽取器（新链路①，替代 parse_pack 做硬约束）。

解决的问题（用户 2026-11-25 拍板重做）：
  parse_pack 在"佳沛...16粒(30-33# 单果克重103-124g)"上把"单果克重103-124g"
  误当整件规格 base=124g，漏掉真正的包装件数 16粒 → 硬约束放过
  "16粒礼盒 vs 猫超无规格散装"的错配。

设计原则：
  1. 件数（count）优先抓「包装件数」：N粒/N个/N盒/N瓶/N袋/N片/N支…
     并排除"单果克重/单片/每粒/单粒"这类描述里的数字。
  2. 规格（base）只认净含量：ml/g/kg/l/斤，且排除"单果克重/单粒重"语境。
  3. 单果/单粒重量单独记为 unit_weight，不进 base（避免与整件规格混淆）。

返回结构：
  {brand_hint, count, count_raw, base, unit, unit_weight, category_hint}
"""
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "pricing_output"))

from build_similarity import full2half, UNIT_WORD  # noqa: E402

# 包装件数量词（件数用）；注意"片/粒"既可能是包装件数也可能是单份描述，需结合语境
PACK_UNITS = ["盒", "瓶", "袋", "桶", "罐", "包", "箱", "支", "只", "枚", "把", "双", "组", "卷", "提"]
ITEM_UNITS = ["粒", "个", "片", "条", "张", "块"]

# 单份重量语境词：出现这些词时，其后的 g/ml 是"单份重量"不是整件规格
UNIT_WEIGHT_CTX = re.compile(r"(单果|单粒|单片|单个|每粒|每片|每个|单颗|每颗|单包|每包)")


def _norm_unit(u):
    u = u.lower()
    if u in ("kg", "千克", "公斤"):
        return 1000.0, "g"
    if u in ("l", "升", "liter"):
        return 1000.0, "ml"
    if u == "斤":
        return 500.0, "g"
    if u in ("毫升",):
        return 1.0, "ml"
    if u in ("克",):
        return 1.0, "g"
    return 1.0, UNIT_WORD.get(u, u)


def extract(name):
    """从品名抽结构属性。所有正则基于半角小写文本。"""
    s = full2half(str(name or "")).lower()
    out = {"brand_hint": "", "count": 1, "count_raw": "", "base": "",
           "unit": "", "unit_weight": "", "category_hint": ""}

    # ---- 1. 单份重量：记录但不进 base ----
    uw = re.search(r"(单果|单粒|单片|单个|每粒|每片|每个|单颗|每颗)\s*克重\s*(\d+(?:\.\d+)?)\s*(g|克|ml|毫升)?", s)
    if uw:
        unit = uw.group(3) or "g"
        out["unit_weight"] = f"{uw.group(2)}{unit}"

    # ---- 2. 件数：优先「N + 包装单位」，排除单份语境 ----
    # 2a. 显式包装件数：N盒/N瓶/N袋/...（含 *N / xN 形式）
    m = re.search(r"(?<!\d)(\d{1,4})\s*(盒|瓶|袋|桶|罐|包|箱|支|只|枚|把|双|组|卷|提)(装)?(?=[\s（(）)、,，。；;]|$)", s)
    if m:
        out["count"], out["count_raw"] = int(m.group(1)), m.group(0)
    else:
        # 2b. 规格*件数：500ml*12 / 200g*8
        m = re.search(r"(\d+(?:\.\d+)?)\s*(ml|g|kg|l|升|克|毫升)\s*[*×x]\s*(\d+)", s)
        if m:
            out["count"], out["count_raw"] = int(m.group(3)), m.group(0)
        else:
            # 2c. 数字+量词：N粒/N个/N片（排除前面紧贴"单/每"的单份语境）
            m = re.search(r"(?<!\d)(\d{1,4})\s*(粒|个|片|条|张|块)(?![\w装])", s)
            if m:
                # 排除"单果克重103-124g"这类：数字 portion 若被 UNIT_WEIGHT_CTX 紧邻覆盖则跳过
                ctx = s[max(0, m.start() - 6):m.end() + 6]
                if not UNIT_WEIGHT_CTX.search(ctx):
                    out["count"], out["count_raw"] = int(m.group(1)), m.group(0)
                else:
                    cm = re.search(r"([双两三四五六])\s*(瓶|袋|盒|罐|支|枚|只)", s)
                    if cm:
                        cn = {"双": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6}
                        out["count"], out["count_raw"] = cn[cm.group(1)], cm.group(0)

    # ---- 3. 规格净含量：排除单份重量语境 ----
    ms = re.search(r"(\d+(?:\.\d+)?)\s*(kg|毫升|ml|mg|g|l|升|斤|克)(?![a-z0-9])", s)
    if ms:
        # 如果匹配落在"单果克重/单粒..."语境里，跳过（那是单份重量）
        ctx_start = max(0, ms.start() - 10)
        ctx = s[ctx_start:ms.end()]
        if not UNIT_WEIGHT_CTX.search(ctx):
            v, u = _norm_unit(ms.group(2))
            val = float(ms.group(1)) * v
            out["base"] = f"{val:g}{u}"
            out["unit"] = u

    return out


def attrs_of(name):
    """供 hard_ok 用：返回是否含明确包装件数、规格base、件数。"""
    a = extract(name)
    return {"pack_explicit": bool(a["count_raw"]) and a["count"] > 1,
            "base": a["base"], "count": a["count"], "count_raw": a["count_raw"],
            "unit_weight": a["unit_weight"]}


if __name__ == "__main__":
    tests = [
        "佳沛新西兰阳光金奇异果礼盒装16粒(30-33# 单果克重103-124g)",
        "佳沛新西兰进口阳光金奇异果黄心新鲜水果送礼孕妇",
        "百菲酪 水牛纯牛奶 200mL*12盒",
        "胡姬花古法小榨花生油400ml",
        "认养一头牛 Hello Kitty A2β-酪蛋白纯牛奶200ml*10盒*2箱",
        "鲁花 5S压榨一级花生油 4L",
        "舒客46孔宽头软毛牙刷8支装",
        "【云仓】卡拉羊行李箱24英寸大容量云朵箱",
    ]
    for t in tests:
        a = extract(t)
        print(f"{t[:44]:46s} → count={a['count']:>3}({a['count_raw'][:10]:10s}) base={a['base']:8s} 单份={a['unit_weight']}")
