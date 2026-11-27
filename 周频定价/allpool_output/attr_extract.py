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
# 营养含量词：出现这些词时，附近的 g/ml 是成分含量（3.8g蛋白、4.3g乳蛋白），不是商品规格
NUTRIENT_WORDS = ("蛋白", "脂肪", "碳水", "钠", "钙", "糖", "能量", "热量", "纤维", "胆固醇", "维生素")


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
    # 2a. 显式包装件数：N盒/N瓶/N袋/...
    # 用户 11-27 拍板修：量词后允许紧跟中文（"8支家庭装颜色随机""106片日夜组合"），
    # 之前尾部约束要求空格/标点，导致这类件数漏判→折算价不除件数虚高。
    # 防误判：①前缀排除范围写法"45-60只"里的 60；②上下文排除单份语境（单果/每片克重）。
    found = None
    for mm in re.finditer(r"(?<![\d\-~～])(\d{1,4})\s*(盒|瓶|袋|桶|罐|包|箱|支|只|枚|把|双|组|卷|提|片|条|粒|个|张|块)", s):
        ctx = s[max(0, mm.start() - 2):mm.end() + 2]
        if UNIT_WEIGHT_CTX.search(ctx):
            continue
        # 排除品牌/型号里的 4 位数字紧贴量词（旭日1996双拼鸭蛋18枚 里的 1996）
        if len(mm.group(1)) == 4 and int(mm.group(1)) >= 1000 and mm.end() < len(s) and s[mm.end()] in "双拼":
            continue
        found = mm
        break
    if found:
        out["count"], out["count_raw"] = int(found.group(1)), found.group(0)
    else:
        # 2b. 规格*件数：500ml*12 / 200g*8
        m = re.search(r"(\d+(?:\.\d+)?)\s*(ml|g|kg|l|升|克|毫升)\s*[*×x]\s*(\d+)", s)
        if m:
            out["count"], out["count_raw"] = int(m.group(3)), m.group(0)
        else:
            # 2c. 中文数字件数：双瓶/两袋/三盒
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


def spec_values(name):
    """抽取清洗后的可比规格数值，供 18 倍差异规则用（用户 2026-11-27 拍板）。

    只保留三类真实可比的规格维度：长度(len)/容量重量(vol)/包装件数(count)。
    清洗规则（避免误伤真同款，实测 69码 2539 对）：
      · 排除营养含量：3.8g蛋白、4.3g乳蛋白、3.3g/100ml（NUTRIENT_WORDS）
      · 排除单份克重：630g(30g*21包) 里的 30g（*N包 语境）
      · 排除型号数字：茶花302003、6类 CAT6、5S压榨
      · 排除范围粒度：45-60只/盒（前带 - ~ ～）
      · 同维度只取一个代表值（最大值），避免单包×N包里的小值混入
    返回 {dim: [值]}，值为统一单位后的浮点数：len→米, vol→毫升/克基数, count→件。
    """
    s = full2half(str(name or "")).lower()
    out = {"len": [], "vol": [], "count": []}

    # 长度：米/cm/mm/英寸 → 统一到米；排除线芯直径(0.008mm)、型号数字
    lens = []
    for v, u in re.findall(r"(\d+(?:\.\d+)?)\s*(米|cm|厘米|mm|毫米|英寸|寸|m)(?![a-z])", s):
        # 排除纯型号数字（茶花302003 M、A0412 等）：6位以上的裸编号
        if len(str(v).replace(".", "")) >= 6 and float(v) > 10000:
            continue
        unit_v = {"米": 1.0, "m": 1.0, "cm": 0.01, "厘米": 0.01, "mm": 0.001,
                  "毫米": 0.001, "英寸": 0.0254, "寸": 0.0333}.get(u, 1.0)
        meters = float(v) * unit_v
        if meters < 0.01:   # 0.008mm 线芯直径这类，不是商品规格
            continue
        lens.append(meters)
    if lens:
        out["len"] = [max(lens)]

    # 容量/重量：ml/l/g/kg/斤 → 统一基数值（ml 与 g 分开单位语义，但同维度内比值可比）
    vols = []
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*(kg|毫升|ml|mg|g|l|升|斤|克)(?![a-z0-9])", s):
        # 营养词：紧跟单位后（3.8g蛋白），或"3.3g/100ml乳蛋白"隔着单位/斜杠的营养标注
        ctx = s[m.start():m.end() + 4]
        ctx_wide = s[m.start():m.end() + 12]
        if (any(w in ctx for w in NUTRIENT_WORDS)
                or ("/" in ctx_wide[:ctx_wide.find("/") + 8] and any(w in ctx_wide for w in NUTRIENT_WORDS))
                or UNIT_WEIGHT_CTX.search(s[max(0, m.start() - 2):m.end() + 2])):
            continue
        # 排除型号/材质里的 L：316L不锈钢、DS-14070L、DL1028L、MK-YSNC1501P
        # 真容量是短数字(2L/1.5L/10L)，型号数字一般≥3位且紧贴中文或处于型号串中
        if m.group(2) == "l":
            raw = m.group(1)
            tail = s[m.end():m.end() + 4]
            if (len(raw) >= 3 and ("不锈钢" in tail or "材质" in tail or "钢" in tail)) or re.match(r"[a-z]", tail):
                continue
            if len(raw) >= 4:  # 14070L / 1028L 型号
                continue
        # 单份克重（30g*21包）不用单独排除：同维度取 max 时 630g 会盖过 30g
        factor, _ = _norm_unit(m.group(2))
        try:
            val = float(m.group(1)) * factor
        except ValueError:
            continue
        # mg 换算后极小，跳过（营养/添加剂量级）
        if m.group(2) == "mg" or val < 1:
            continue
        vols.append(val)
    if vols:
        out["vol"] = [max(vols)]

    # 包装件数：复用 extract 的 count，仅 >1 才算（1 件不构成倍数比较）
    # 但"1瓶 vs 24瓶"这类是用户明确要判的场景：两边都写到件数维度时 1 也要进集合，
    # 由调用方仅在一边有件数另一边无件数时跳过该维度比较。
    a = extract(name)
    if a["count"] >= 1 and a["count_raw"]:
        out["count"] = [float(a["count"])]
    return out


def attrs_of(name):
    """供 hard_ok 用：返回是否含明确包装件数、规格base、件数。"""
    a = extract(name)
    return {"pack_explicit": bool(a["count_raw"]) and a["count"] > 1,
            "base": a["base"], "count": a["count"], "count_raw": a["count_raw"],
            "unit_weight": a["unit_weight"], "specs": spec_values(name)}


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
