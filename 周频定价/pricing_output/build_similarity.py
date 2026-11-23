#!/usr/bin/env python3
"""构建「同品关系」两层清单：
  第一层 = 历史比价表已人工标「同品同规格=是」的禄品P码↔京东SKU（直接采信，附京东市场价比对参考）
  第二层 = 其余禄品SKU，用名称相似度算候选京东品 + 相似度% + 质疑点，供人工确认。绝不自动判定同款。

只读源文件；路径由 __file__ 推导，env 可覆盖：BENLINK_REPO / PRICING_DIR。
"""
import csv
import os
import re
import json
import difflib
from collections import defaultdict, Counter
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(os.environ.get("PRICING_DIR", HERE))
REPO = Path(os.environ.get("BENLINK_REPO", Path.home() / "ai-workspace" / "projects" / "benlink-product"))

BASE_XLSX = REPO / "01_源数据" / "禄品4940_筛选表_2026-11-09.xlsx"
HIST_XLSX = PRICING_DIR / "源数据_京东比价总表1.29.xlsx"
VOP_CSV = HERE.parent / "supply_pull" / "output" / "本来生活VOP_goods_本网ID.csv"
PRICE_XLSX = PRICING_DIR / "规则定价_成本与京东比价_2026-11-13.xlsx"
OUTPUT = PRICING_DIR / "同品关系_候选相似度与人工确认_2026-11-16.xlsx"

# 相似度权重（可解释、可调整）
# 规则1 品牌+相似度并重；规则2/3/4 规格强一致(单位归一+数量换算)；规则5 双向价格反证
W_SEQ = 0.22          # 字符序列相似度
W_BIGRAM = 0.22       # 二字词 Jaccard
BONUS_BRAND = 0.24    # 品牌一致加分
PENALTY_BRAND = -0.25 # 品牌不一致扣分
BONUS_MODEL = 0.14    # 型号一致加分
PENALTY_MODEL = -0.08 # 双方型号都不同扣分
BONUS_SPEC = 0.18     # 规则2/3：规格基数强一致加分（归一化后同规格=强信号）
PENALTY_SPEC = -0.22  # 规格基数不一致强扣分（不同规格多半不是同一个品）
PENALTY_PRICE = -0.15 # 规则5：双向单价差异过大时降分（数量已换算后仍差很多）
PENALTY_CATEGORY = -0.55  # 品类冲突（稻米油 vs 牛奶）：强扣，基本排除
CAND_THRESHOLD = 0.40 # 进入候选清单的最低分（真同款名字差异大，故调低，交人工判断）
TOP_N = 5             # 每个禄品最多保留候选数
PRICE_RATIO_HI = 2.5  # 规则5：单位价之比 >2.5 或 <1/2.5 视为差异过大，降分+质疑
PRICE_RATIO_LO = 1/2.5

# 品类词典（识别商品类型，防止"油 vs 牛奶"这类跨品类误判）
# 顺序：具体词在前；不收"油/奶"这类单字，避免"无硅油""控油"误命中
CATEGORY_TERMS = [
    ("食用油", ["稻米油", "米糠油", "菜籽油", "花生油", "葵花籽油", "玉米油", "橄榄油",
              "调和油", "大豆油", "食用油", "色拉油", "芝麻油", "香油", "亚麻籽油"]),
    ("乳制品", ["纯牛奶", "鲜牛奶", "纯奶", "酸奶", "牛奶", "奶酪", "奶昔", "乳蛋白", "调制乳"]),
    ("洗发", ["洗发水", "洗发露", "洗发液", "洗头水"]),
    ("护发", ["护发素", "护发乳", "发膜", "润发乳"]),
    ("沐浴", ["沐浴露", "沐浴乳", "沐浴盐", "香皂"]),
    ("益生菌", ["益生菌", "乳酸菌", "益生元"]),
    ("膳食补充", ["膳食纤维", "维生素", "白芸豆", "胶原蛋白", "蛋白粉", "叶黄素", "钙片"]),
    ("代餐", ["代餐", "能量棒", "蛋白棒"]),
    ("酒类", ["白酒", "红酒", "葡萄酒", "啤酒", "黄酒", "米酒", "洋酒", "威士忌"]),
    ("饮料", ["果汁", "可乐", "气泡水", "茶饮料", "咖啡", "饮用水", "矿泉水"]),
    ("米面杂粮", ["小米", "大米", "面粉", "杂粮", "糙米", "燕麦", "挂面", "意大利面"]),
    ("水果", ["苹果", "蓝莓", "橙子", "香蕉", "葡萄", "草莓", "车厘子", "猕猴桃"]),
    ("纸品", ["纸巾", "卷纸", "抽纸", "湿巾", "卫生纸"]),
    ("家清", ["洗衣液", "洗洁精", "消毒液", "柔顺剂", "清洁剂"]),
    ("食用油" , ["菜油", "猪油"]),
]


def category_of(name):
    """从商品名提取品类标签集合（可多标签，如'洗护套装'）。"""
    s = full2half(str(name or "")).lower()
    tags = set()
    for cat, terms in CATEGORY_TERMS:
        if any(t in s for t in terms):
            tags.add(cat)
    return tags

MKT_WORDS = ["正品", "旗舰店", "旗舰", "官方", "自营", "进口", "原装", "全新", "升级版", "加强版",
             "家用", "便携", "送礼", "礼品", "批发", "特价", "包邮", "精选", "优选", "尊享", "巅峰",
             "组合装", "家庭装", "促销", "爆款", "热卖", "新一代", "经典"]

C_HEADER = "1F4E79"; C_BASE = "D9E2F3"; C_CONF = "C6EFCE"; C_CAND = "FFF2CC"
C_HIGH = "C6EFCE"; C_MID = "FFEB9C"; C_LOW = "FCE4D6"; C_WARN = "FFC7CE"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def full2half(s):
    out = []
    for ch in s:
        code = ord(ch)
        if code == 0x3000:
            out.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out)


def num(v):
    try:
        x = float(str(v).strip())
        return x if x > 0 else None
    except Exception:
        return None


def brand_main(s):
    """从 VOP brand_name 提取中文主品牌（去括号英文、去商店后缀）。"""
    s = full2half(str(s or ""))
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    s = re.sub(r"(官方|旗舰|自营|专营|专卖|店)$", "", s.strip())
    s = re.sub(r"[A-Za-z0-9\s]+", "", s)
    return s.strip()


def brand_tokens(s):
    """从商品名开头提取品牌的多粒度token（2/3/4字），用于品牌一致性判断。"""
    s = re.sub(r"【[^】]*】", "", str(s or "")).strip()
    s = re.sub(r"^[（(]?[A-Za-z]+[）)]?\s*", "", s)
    m = re.match(r"^([\u4e00-\u9fa5]{2,6})", s)
    if m:
        w = m.group(1)
        return {w[:2], w[:3], w[:4]}
    return set()


def model_codes(s, brand_alias=""):
    """提取型号集合（含字母数字串），剔除纯规格(ml/g)与品牌英文名。"""
    toks = re.findall(r"[A-Za-z0-9][A-Za-z0-9.\-]{2,}", full2half(str(s or "")))
    stop = {"pro", "plus", "max", "mini", "new", "the", "for", "with", "set", "and",
            "version", "upgrade", "gift", "box", "ml", "kg", "pcs", "glass", "bag", "bottle"}
    alias = set(re.findall(r"[A-Za-z]+", brand_alias or ""))
    alias |= {"midea", "braun", "sisley", "ecovacs", "philips", "panasonic", "mouton", "chateau"}
    out = set()
    for t in toks:
        tl = t.lower()
        if tl in stop or tl in alias or len(tl) < 3 or not re.search(r"[A-Za-z]", t):
            continue
        if re.fullmatch(r"\d+(ml|g|kg)", tl):
            continue
        out.add(tl)
    return out


def model_code(s, brand_alias=""):
    cs = model_codes(s, brand_alias)
    return max(cs, key=len) if cs else ""


def spec_norm(s):
    """提取并归一化净含量（g / ml / kg / L），用于规格一致性判断。"""
    s = full2half(str(s or "")).lower()
    m = re.search(r"(\d+(?:\.\d+)?)\s*(kg|g|ml|l)(?!\w)", s)
    if not m:
        return ""
    v = float(m.group(1)); u = m.group(2)
    if u == "kg": return f"{v*1000:.0f}g"
    if u == "l":  return f"{v*1000:.0f}ml"
    return f"{v:g}{u}"


# ===== 规则2/3/4：规格 / 单位 / 数量 解析 =====
UNIT_WORD = {"克": "g", "kg": "g", "千克": "g", "公斤": "g", "g": "g",
             "毫升": "ml", "ml": "ml", "l": "ml", "升": "ml",
             "斤": "jin", "片": "片", "粒": "粒", "支": "支", "袋": "袋", "个": "个"}


def parse_pack(name):
    """从名称解析 (规格基数归一串, 件数n, 明细串)。
    规则2: 规格在名称里; 规则3: g/克、ml/毫升、kg/升 归一; 规则4: *2 / 500ml+500ml / 2瓶装 识别件数。
    返回 spec=('{数值}{单位}' 归一), count=int, raw=原始描述。
    """
    s = full2half(str(name or "")).lower()
    # 件数：紧贴规格的 *N / ×N / xN
    mult = 1
    m = re.search(r"(\d+(?:\.\d+)?)\s*(kg|ml|g|l|斤|片|粒|支|袋|个|盒|瓶|罐|升|毫升|克)\s*[*×x]\s*(\d+)", s)
    if m:
        mult = int(m.group(3))
        core = (m.group(1), m.group(2)); raw = f"{m.group(1)}{m.group(2)}*{m.group(3)}"
    else:
        # 500ml+500ml / 500g+500g 之类叠加
        m2 = re.search(r"(\d+(?:\.\d+)?)\s*(kg|ml|g|l|克|毫升|斤)\s*\+\s*(\d+(?:\.\d+)?)\s*(kg|ml|g|l|克|毫升|斤)", s)
        if m2 and abs(float(m2.group(1)) - float(m2.group(3))) < 1e-6 and UNIT_WORD.get(m2.group(2)) == UNIT_WORD.get(m2.group(4)):
            mult = 2; core = (m2.group(1), m2.group(2)); raw = f"{m2.group(1)}{m2.group(2)}+{m2.group(3)}{m2.group(4)}"
        else:
            # 2瓶装 / 3袋装 / 双支 等
            mc = re.search(r"(?<!\d)(\d+)\s*(瓶|袋|盒|罐|支|枚|只)\s*装", s)
            cn = {"双": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6}
            mcn = re.search(r"([双两三四五六])\s*(瓶|袋|盒|罐|支|枚|只)", s)
            if mc:
                mult = int(mc.group(1)); raw = mc.group(0)
            elif mcn:
                mult = cn[mcn.group(1)]; raw = mcn.group(0)
            else:
                raw = ""
            # 规格基数（主净含量，取第一个）
            ms = re.search(r"(\d+(?:\.\d+)?)\s*(kg|毫升|ml|mg|g|l|升|斤|克)(?![\w])", s)
            if not ms:
                # 无净含量时，尝试纯件数（京东名常写"180瓶""30粒""60片"）
                mc2 = re.search(r"(?<!\d)(\d{1,4})\s*(瓶|粒|片|支|袋|盒|罐|枚|只|桶)(?![\w装])", s)
                if mc2:
                    mult = int(mc2.group(1)); raw = mc2.group(0)
                return {"base": "", "count": mult, "total": "", "raw": raw, "unit": "", "value": None}
            core = (ms.group(1), ms.group(2))
            if not raw:
                raw = f"{ms.group(1)}{ms.group(2)}"
    v = float(core[0]); u = core[1]
    # 规则3：单位归一
    if u in ("kg", "千克", "公斤"):
        v *= 1000; u = "g"
    elif u == "升" or u == "l":
        v *= 1000; u = "ml"
    elif u == "斤":
        v *= 500; u = "g"
    else:
        u = UNIT_WORD.get(u, u)
    spec = f"{v:g}{u}"
    # 规则4：把件数折算进「总净含量」用于强一致比较（500ml*2 = 1000ml）
    total_v = v * mult
    return {"base": spec, "count": mult, "total": f"{total_v:g}{u}", "raw": raw, "unit": u, "value": v}


def pack_of(name, extra_spec=""):
    """优先用规格字段，其次名称；返回 parse_pack 结构。"""
    p = parse_pack(name)
    e = parse_pack(extra_spec)
    if (not p or not p.get("base")) and e.get("base"):
        return e
    if e.get("base") and p.get("base") and e["total"] != p["total"]:
        # 名称与规格字段都有但不一致时，以名称为准（名称更贴近展示全称）
        return p
    return p


def unit_price(price, pack):
    """规则5：折算成单位净含量价（元/基础规格），用于跨数量比价。"""
    if not price or not pack or not pack.get("value"):
        return None
    return price / (pack["value"] * pack["count"])


def normalize(s, strip_spec=False):
    s = full2half(str(s or "")).lower()
    s = re.sub(r"【[^】]*】", "", s)
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    for w in MKT_WORDS:
        s = s.replace(w.lower(), "")
    if strip_spec:
        s = re.sub(r"\d+(?:\.\d+)?\s*(kg|g|ml|l|斤|片|粒|个|支)", "", s)
    s = re.sub(r"[\s\-—_/·,，。.+*&'\"’“”()（）\[\]【】!！:：;；|]+", "", s)
    return s


def bigrams(s):
    return {s[i:i+2] for i in range(len(s) - 1)} if len(s) >= 2 else ({s} if s else set())


def score_pair(lp_name, vop_name, vop_brand="", lp_model="", vop_model="",
               lp_pack=None, vop_pack=None, our_price=None, jd_price=None):
    """返回 (总分0-1, 品牌核对, 型号核对, 规格核对, 价格核对, 质疑点列表)。

    规则1 品牌+相似度；规则2/3/4 规格强一致(单位归一+数量换算)；规则5 双向价格反证。
    lp_pack/vop_pack 为 parse_pack 结构；our_price/jd_price 用于单价反证。
    """
    n1 = normalize(lp_name, strip_spec=False)
    n2 = normalize(vop_name, strip_spec=False)
    seq = difflib.SequenceMatcher(None, n1, n2).ratio()
    b1, b2 = bigrams(normalize(lp_name, strip_spec=True)), bigrams(normalize(vop_name, strip_spec=True))
    jac = len(b1 & b2) / len(b1 | b2) if (b1 | b2) else 0.0

    # 规则1：品牌（VOP 给 brand_name，禄品从名称开头提取）
    bt_lp = brand_tokens(lp_name)
    bt_vop = brand_tokens(vop_name)
    bm = brand_main(vop_brand)
    if bm:
        bt_vop |= {bm[i:i+k] for k in (2, 3, 4) for i in range(max(1, len(bm) - k + 1))} | {bm[:2], bm[:3], bm[:4]}
    brand_ok = bool(bt_lp & bt_vop)
    brand_flag = "一致" if brand_ok else ("不一致" if bt_lp and bt_vop else "无法判断")
    brand_adj = BONUS_BRAND if brand_ok else (PENALTY_BRAND if bt_lp and bt_vop else 0.0)

    # 型号：集合交集
    ml = lp_model if isinstance(lp_model, set) else ({lp_model} if lp_model else set())
    mv = vop_model if isinstance(vop_model, set) else ({vop_model} if vop_model else set())
    inter = {x for x in ml & mv if len(x) >= 3}
    if inter:
        model_adj, model_flag = BONUS_MODEL, f"一致({'/'.join(sorted(inter))})"
    elif ml and mv:
        model_adj, model_flag = PENALTY_MODEL, f"不同({'/'.join(sorted(ml))[:20]} vs {'/'.join(sorted(mv))[:20]})"
    elif not ml and not mv:
        model_adj, model_flag = 0.0, "均无型号"
    else:
        model_adj, model_flag = 0.0, "一方缺型号"

    # 规则2/3/4：规格强一致 —— 先比「基础规格base」（同款必须单件规格相同），再看件数
    # 关键修正：总净含量相同但基础规格不同(如 3L单瓶 vs 250ml*12)绝不算同规格
    lb = lp_pack.get("base") if lp_pack else ""
    vb = vop_pack.get("base") if vop_pack else ""
    lc = lp_pack.get("count", 1) if lp_pack else 1
    vc = vop_pack.get("count", 1) if vop_pack else 1
    lt = lp_pack.get("total") if lp_pack else ""
    vt = vop_pack.get("total") if vop_pack else ""
    if lb and vb:
        if lb == vb:
            # 基础规格相同：再看件数
            if lc == vc:
                spec_adj, spec_flag = BONUS_SPEC, f"规格一致({lb}×{lc})"
            else:
                # 规则4：基础规格同、件数不同 = 同款不同包装（Spes 500ml*2 vs 500ml）
                spec_adj, spec_flag = BONUS_SPEC * 0.5, f"同规格不同件数({lb} 件数{lc}vs{vc})"
        else:
            # 基础规格不同——即便总净含量碰巧相等(3L vs 250ml*12)也判规格不同
            note = " 总量同但单件规格异" if lt == vt else ""
            spec_adj, spec_flag = PENALTY_SPEC, f"规格不同({lb} vs {vb}){note}"
    else:
        spec_adj, spec_flag = 0.0, f"规格({lb or '无'}/{vb or '无'})"

    # 品类冲突（稻米油 vs 牛奶）：不同品类基本可排除
    cat_lp = category_of(lp_name)
    cat_vp = category_of(vop_name)
    if cat_lp and cat_vp and not (cat_lp & cat_vp):
        cat_adj = PENALTY_CATEGORY
        cat_flag = f"品类冲突({'/'.join(sorted(cat_lp))} vs {'/'.join(sorted(cat_vp))})"
    else:
        cat_adj = 0.0
        cat_flag = f"品类一致({'/'.join(sorted(cat_lp & cat_vp))})" if (cat_lp & cat_vp) else "品类未识别"

    # 规则5：双向价格反证 —— 用单位净含量价比较，跨数量可比；差异过大降分
    pup = unit_price(our_price, lp_pack)
    pjd = unit_price(jd_price, vop_pack)
    price_adj = 0.0
    if pup and pjd and pjd > 0:
        ratio = pup / pjd  # 我方单位价 / 京东单位价（双向：过小=我们便宜太多，过大=贵太多）
        if ratio > PRICE_RATIO_HI or ratio < PRICE_RATIO_LO:
            price_adj = PENALTY_PRICE
            price_flag = f"单价差大({ratio:.1f}x)"
        else:
            price_flag = f"单价接近({ratio:.2f}x)"
    else:
        price_flag = "无价可核"

    total = W_SEQ * seq + W_BIGRAM * jac + brand_adj + model_adj + spec_adj + price_adj + cat_adj
    total = max(0.0, min(1.0, total))

    doubts = []
    if cat_adj:
        doubts.append(f"品类不同({'/'.join(sorted(cat_lp))}/{'/'.join(sorted(cat_vp))}),基本可排除")
    if not brand_ok and bt_lp and bt_vop:
        doubts.append(f"品牌不一致(京东:{bm or '/'.join(sorted(bt_vop))[:8]})")
    if ml and mv and not inter:
        doubts.append("型号不同")
    if lb and vb:
        if lb != vb:
            doubts.append(f"规格不同({lb} vs {vb})" + ("·总量同但单件异" if lt == vt else ""))
        elif lc != vc:
            doubts.append(f"同规格不同件数({lc}vs{vc})·核实包装")
    if price_adj:
        doubts.append(f"单价差大(我方/京东={pup/pjd:.1f}x),疑非同款" if pup and pjd else "单价差大")
    if total < 0.70:
        doubts.append("名称相似度偏低,需重点核")

    # 统一优先级：告诉用户先看什么、次看什么
    level, focus = screening_priority(cat_adj, brand_ok, lb, vb, lc, vc, ml, mv, inter, price_adj, seq)
    return total, brand_flag, model_flag, spec_flag, price_flag, cat_flag, level, focus, doubts


def screening_priority(cat_adj, brand_ok, lb, vb, lc, vc, ml, mv, inter, price_adj, seq):
    """给每条候选一个统一的复核优先级 + 关注点，指导用户先看什么、次看什么。
    返回 (level, focus)。level: ①优先 / ②次看 / ③存疑 / ④排除
    """
    # 硬性排除：品类冲突
    if cat_adj:
        return "④排除", "品类不同，直接排除"
    # 规格不同（基础规格不同）→ 基本排除，除非型号强证据
    spec_diff = bool(lb and vb and lb != vb)
    if spec_diff and not inter:
        return "④排除", "单件规格不同且型号不符，基本可排除"
    if spec_diff and inter:
        return "③存疑", "单件规格不同，仅型号一致；重点核包装规格"

    # 到这一层：规格一致 / 同规格不同件数 / 无规格
    same_pack = bool(lb and vb and lb == vb)
    diff_count = same_pack and lc != vc
    # 品牌不一致
    if not brand_ok and (ml or mv):
        return "③存疑", "品牌存疑；先核品牌，再看型号/规格"

    # 单件规格未知但件数可比（京东常只写"30瓶"不写"2g"）
    spec_unknown = not (lb and vb)
    count_match = (lc == vc)
    # 核心关注点排序：价格反证 > 包装件数 > 型号 > 名称
    if price_adj:
        return "②次看", "单件价差大；先核价格是否同口径，再看包装/规格"
    if diff_count:
        return "②次看", f"同规格不同件数({lc}vs{vc})；先核包装数量，再看单价"
    if inter and same_pack:
        return "①优先", "品牌+型号+规格一致；重点核图片/详情页确认同款"
    if inter and spec_unknown and count_match:
        return "②次看", f"型号一致、件数一致({lc})；京东未写单件规格，先核图片/详情页"
    if inter and spec_unknown and not count_match:
        return "③存疑", f"型号一致但件数不同({lc}vs{vc})；先核包装规格"
    if inter:
        return "②次看", "品牌+型号一致；规格需人工核，重点看图片"
    if same_pack:
        return "①优先", "品牌+规格一致；重点核型号与图片"
    if spec_unknown and count_match:
        return "②次看", f"件数一致({lc})但规格未知；先核京东单件规格与图片"
    # 无规格可比：靠名称
    if seq >= 0.55:
        return "②次看", "无规格字段；名称较像，重点核规格参数与图片"
    return "③存疑", "规格缺失且名称相似度低；谨慎，需人工逐一核"


def load_vop():
    """sku_id -> {name, brand, price, model, spec, bl_code}，jd_price 取最低。"""
    prices = defaultdict(list); info = {}
    with open(VOP_CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            sku = str(row.get("sku_id", "")).strip()
            if not sku:
                continue
            p = num(row.get("jd_price"))
            if p is not None:
                prices[sku].append(p)
            if sku not in info:
                sname = (row.get("sku_name") or "").strip()
                sbrand = (row.get("brand_name") or "").strip()
                info[sku] = {
                    "name": sname,
                    "brand": sbrand,
                    "model": model_codes(sname, sbrand),
                    "spec": spec_norm(sname),
                    "pack": pack_of(sname),
                    "bl_code": (row.get("bl_product_code") or "").strip(),
                    "captured": (row.get("captured_at") or "").strip(),
                }
    for sku, v in info.items():
        v["price"] = min(prices[sku]) if prices.get(sku) else None
    return info


def load_baseline():
    wb = openpyxl.load_workbook(BASE_XLSX, read_only=True, data_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    hdr = list(rows[1])
    out = {}
    for r in rows[2:]:
        d = dict(zip(hdr, r))
        p = str(d.get("商品编号") or "").strip()
        if p:
            out[p] = d
    wb.close()
    return out


def load_hist():
    """禄品P码 -> list of 确认/历史记录。"""
    wb = openpyxl.load_workbook(HIST_XLSX, read_only=True, data_only=True)
    rows = wb.worksheets[0].iter_rows(values_only=True)
    hdr = list(next(rows))
    I = {k: hdr.index(k) for k in ["商品编码", "商品名称", "京东商品名称", "京东规格", "本网规格",
                                  "是否同品同规格", "京东SKU", "京东当前售价", "京东折合价", "更新时间"]}
    out = defaultdict(list)
    for r in rows:
        if not r or r[I["商品编码"]] in (None, ""):
            continue
        p = str(r[I["商品编码"]]).strip()
        out[p].append({
            "sku": str(r[I["京东SKU"]]).strip() if r[I["京东SKU"]] is not None else "",
            "same": str(r[I["是否同品同规格"]]).strip() if r[I["是否同品同规格"]] is not None else "",
            "jname": str(r[I["京东商品名称"]]).strip() if r[I["京东商品名称"]] not in (None, "/") else "",
            "jpack": pack_of(str(r[I["京东商品名称"]] or ""), str(r[I["京东规格"]] or "")),
            "price": num(r[I["京东当前售价"]]) or num(r[I["京东折合价"]]),
            "updated": str(r[I["更新时间"]]).strip() if r[I["更新时间"]] is not None else "",
        })
    wb.close()
    return out


def load_our_price():
    """P码 -> (采用成本, 推荐输出价)。"""
    wb = openpyxl.load_workbook(PRICE_XLSX, read_only=True, data_only=True)
    ws = wb["05_逐SKU定价明细"]
    rows = ws.iter_rows(values_only=True)
    hdr = list(next(rows))
    ip, ic, io = hdr.index("商品编号(P码)"), hdr.index("采用成本(元)"), hdr.index("推荐B_输出价")
    out = {}
    for r in rows:
        p = str(r[ip]).strip()
        out[p] = (num(r[ic]), num(r[io]))
    wb.close()
    return out


def style_header(ws, row, fills, height=34):
    for col, fill in enumerate(fills, start=1):
        c = ws.cell(row=row, column=col)
        c.fill = PatternFill("solid", fgColor=fill)
        c.font = Font(bold=True, color="FFFFFF" if fill == C_HEADER else "1F1F1F", size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = height


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def main():
    for p in (BASE_XLSX, HIST_XLSX, VOP_CSV, PRICE_XLSX):
        if not p.exists():
            raise SystemExit(f"缺少源文件：{p}")

    baseline = load_baseline()
    vop = load_vop()
    hist = load_hist()
    our = load_our_price()
    print(f"载入：禄品 {len(baseline)} / VOP品 {len(vop)} / 历史表P码 {len(hist)} / 定价明细 {len(our)}", flush=True)

    # ---------- 第一层：历史已确认同品同规格 ----------
    layer1 = []
    confirmed_pcodes = set()
    for p in sorted(baseline):
        for rec in hist.get(p, []):
            if rec["same"] != "是" or not rec["sku"]:
                continue
            confirmed_pcodes.add(p)
            sku = rec["sku"]
            vk = vop.get(sku, {})
            jd_price = vk.get("price") or rec["price"]
            jd_src = "VOP实时" if vk.get("price") else ("历史表价" if rec["price"] else "无价")
            cost, our_price = our.get(p, (None, None))
            gap = (our_price - jd_price) if (our_price is not None and jd_price) else None
            gap_pct = (gap / jd_price) if (gap is not None and jd_price) else None
            lp_name = str(baseline[p].get("商品名称") or "")
            lp_pack = pack_of(lp_name)
            vp_pack = vk.get("pack") or rec.get("jpack")
            # 规则4/5：单价（折算每基础规格）与件数比
            pup = unit_price(our_price, lp_pack)
            pjd = unit_price(jd_price, vp_pack)
            unit_ratio = (pup / pjd) if (pup and pjd and pjd > 0) else None
            layer1.append({
                "p": p, "name": lp_name,
                "cat": str(baseline[p].get("二级类目") or ""),
                "sku": sku, "jname": vk.get("name") or rec["jname"],
                "jd_price": jd_price, "jd_src": jd_src,
                "vd_code": vk.get("bl_code", ""), "captured": vk.get("captured", ""),
                "updated": rec["updated"], "cost": cost, "our": our_price,
                "gap": gap, "gap_pct": gap_pct,
                "lp_pack": lp_pack, "vp_pack": vp_pack,
                "lp_spec": (lp_pack or {}).get("base", ""), "vp_spec": (vp_pack or {}).get("base", ""),
                "lp_cnt": (lp_pack or {}).get("count", 1), "vp_cnt": (vp_pack or {}).get("count", 1),
                "unit_ratio": unit_ratio,
            })
    print(f"第一层：已确认同品关系 {len(layer1)} 条，涉及禄品P码 {len(confirmed_pcodes)} 个", flush=True)

    # ---------- 校准：用第一层(禄品名 vs VOP京东名)看打分分布 ----------
    calib = []
    for row in layer1:
        if not row["jname"]:
            continue
        s, *_ = score_pair(row["name"], row["jname"],
                           lp_pack=row["lp_pack"], vop_pack=row["vp_pack"],
                           our_price=row["our"], jd_price=row["jd_price"])
        calib.append(s)
    if calib:
        calib.sort()
        n = len(calib)
        print(f"校准(已确认同款 {n} 对)相似度分布: "
              f"中位 {calib[n//2]:.0%} / >=70% {sum(1 for x in calib if x>=0.7)/n:.0%} / "
              f">=55% {sum(1 for x in calib if x>=0.55)/n:.0%} / <40% {sum(1 for x in calib if x<0.4)/n:.0%}", flush=True)

    # ---------- 第二层：其余禄品，两路候选 ----------
    remaining = [p for p in baseline if p not in confirmed_pcodes]
    # 路A：历史表里该禄品的"否//"京东对照行（现成候选，当年人工判不像，交用户重判）
    # 路B：VOP 实时品按名称相似度召回
    inv = defaultdict(list)
    for sku, v in vop.items():
        if not v["name"] or not v["price"]:   # 只用有京东市场价的品做候选
            continue
        for bg in bigrams(normalize(v["name"], strip_spec=True)):
            inv[bg].append(sku)

    layer2 = []
    no_candidate = 0
    for p in remaining:
        lp = baseline[p]
        lname = str(lp.get("商品名称") or "")
        if not lname:
            continue
        lp_model = model_codes(lname)
        lp_pack = pack_of(lname)
        our_price = our.get(p, (None, None))[1]
        scored = []

        # 路A：历史表该P码的非"是"对照行
        for rec in hist.get(p, []):
            if rec["same"] == "是" or not rec["jname"]:
                continue
            s, bf, mf, sf, pf, cf, level, focus, doubts = score_pair(lname, rec["jname"], "", lp_model, "",
                                                   lp_pack, rec.get("jpack"), our_price, rec["price"])
            if s >= CAND_THRESHOLD:
                scored.append(dict(s=s, sku=rec["sku"], jname=rec["jname"], brand="",
                                   bf=bf, mf=mf, sf=sf, pf=pf, cf=cf, level=level, focus=focus, doubts=doubts,
                                   jd_price=rec["price"], captured="", bl_code="",
                                   lp_spec=(lp_pack or {}).get("base", ""), vp_spec=(rec.get("jpack") or {}).get("base", ""),
                                   lp_cnt=(lp_pack or {}).get("count", 1), vp_cnt=(rec.get("jpack") or {}).get("count", 1),
                                   source="历史表对照", hist_time=rec["updated"]))

        # 路B：VOP 相似度召回
        cand = Counter()
        for bg in bigrams(normalize(lname, strip_spec=True)):
            for sku in inv.get(bg, []):
                cand[sku] += 1
        for sku, _ in cand.most_common(60):
            v = vop[sku]
            s, bf, mf, sf, pf, cf, level, focus, doubts = score_pair(lname, v["name"], v["brand"], lp_model, v["model"],
                                                   lp_pack, v["pack"], our_price, v["price"])
            if s >= CAND_THRESHOLD:
                scored.append(dict(s=s, sku=sku, jname=v["name"], brand=v["brand"],
                                   bf=bf, mf=mf, sf=sf, pf=pf, cf=cf, level=level, focus=focus, doubts=doubts,
                                   jd_price=v["price"], captured=v["captured"], bl_code=v["bl_code"],
                                   lp_spec=(lp_pack or {}).get("base", ""), vp_spec=(v["pack"] or {}).get("base", ""),
                                   lp_cnt=(lp_pack or {}).get("count", 1), vp_cnt=(v["pack"] or {}).get("count", 1),
                                   source="VOP实时", hist_time=""))

        # 同一SKU去重，保留最高分
        best = {}
        for x in scored:
            k = x["sku"] or x["jname"]
            if k not in best or x["s"] > best[k]["s"]:
                best[k] = x
        # 排序：先按复核优先级(①→④)，同级内按【规格相似度】优先（用户要求），再按分数
        # 规格相似度：单件规格+件数全同 > 同规格不同件数 > 规格未知但件数接近 > 规格未知件数差异大 > 规格不同
        def spec_rank(x):
            ls, vs = x.get("lp_spec", ""), x.get("vp_spec", "")
            lc, vc = x.get("lp_cnt", 1), x.get("vp_cnt", 1)
            if ls and vs and ls == vs and lc == vc:
                return (0, 0)          # 规格+件数完全一致
            if ls and vs and ls == vs:
                return (1, 0)          # 同规格不同件数
            if not ls or not vs:
                # 一方规格未知 → 用件数接近度决定（30瓶 vs 30瓶 优先于 30瓶 vs 180瓶）
                try:
                    gap = abs(float(lc) - float(vc))
                except Exception:
                    gap = 999
                if lc == vc:
                    return (2, 0)      # 件数一致
                return (3, min(gap, 1000))  # 件数差越小越靠前
            return (5, 0)              # 规格不同
        level_order = {"①优先": 0, "②次看": 1, "③存疑": 2, "④排除": 3}
        scored = sorted(best.values(),
                        key=lambda x: (level_order.get(x["level"], 9), spec_rank(x), -x["s"]))
        if not scored:
            no_candidate += 1
            continue
        for rank, x in enumerate(scored[:TOP_N], start=1):
            cost, our_price = our.get(p, (None, None))
            layer2.append({
                "p": p, "name": lname, "cat": str(lp.get("二级类目") or ""),
                "rank": rank, "source": x["source"], "sku": x["sku"], "jname": x["jname"], "brand": x["brand"],
                "sim": x["s"], "brand_flag": x["bf"], "model_flag": x["mf"],
                "spec_flag": x["sf"], "price_flag": x["pf"], "cat_flag": x["cf"],
                "level": x["level"], "focus": x["focus"],
                "lp_spec": x["lp_spec"], "vp_spec": x["vp_spec"],
                "lp_cnt": x["lp_cnt"], "vp_cnt": x["vp_cnt"],
                "doubts": "；".join(x["doubts"]) if x["doubts"] else "无明显质疑点",
                "jd_price": x["jd_price"], "captured": x["captured"], "hist_time": x["hist_time"],
                "cost": cost, "our": our_price, "bl_code": x["bl_code"],
                "confirm": "", "note": "",
            })
    print(f"第二层：待确认候选 {len(layer2)} 条（{len(remaining)} 个禄品中 {no_candidate} 个无候选）", flush=True)

    build_workbook(layer1, layer2, baseline, calib, confirmed_pcodes, remaining)
    print(f"已生成：{OUTPUT}", flush=True)


def build_workbook(l1, l2, baseline, calib, confirmed_pcodes, remaining):
    wb = openpyxl.Workbook(); wb.remove(wb.active)

    # ---- 00 说明 ----
    ws = wb.create_sheet("00_说明")
    set_widths(ws, [4, 26, 104])
    ws["A1"] = "同品关系清单 · 已确认 + 相似度候选（供人工确认）"
    ws["A1"].font = Font(bold=True, size=16, color=C_HEADER); ws.merge_cells("A1:C1")
    ws.row_dimensions[1].height = 30
    lines = [
        ("SEC", "一、为什么有两套 P 码", ""),
        ("TXT", "核心事实", "京东 VOP 文件里每个京东品都带一个本网 P 码（约 4.3 万个），与禄品 4,940 个 P 码交集为 0——确实是两套编码。同一个实物商品，禄品体系一个码、京东体系另一个码。"),
        ("TXT", "证据", "历史比价表中有 11,395 行被人工标为「同品同规格=是」，即禄品 P 码与京东 SKU 被判定为同一个商品。这就是“两个码一个品”的现成证明。"),
        ("SEC", "二、两层清单", ""),
        ("TXT", "第一层 已确认同品", f"历史表标「同品同规格=是」的直接采信，共 {len(l1)} 条、涉及 {len(confirmed_pcodes)} 个禄品 P 码。附京东市场售卖价（VOP 实时价优先，缺失用历史表价）作为我们定价后的比对参考。"),
        ("TXT", "第二层 待确认候选", f"其余 {len(remaining)} 个禄品，从两路找候选：①历史表里当年标“否//”的京东对照（人工曾判不像，请重判）；②VOP 实时品按名称相似度召回。共 {len(l2)} 条。相似度%与质疑点仅作参考，绝不自动判为同款——需您在「人工确认」列填“是/否”。"),
        ("SEC", "三、按您的规则怎么判（可解释）", ""),
        ("TXT", "①品牌+相似度并重", "品牌一致(+24%)/不一致(−25%)是强信号，同时算字符序列(22%)+二字词重合(22%)——两个都看，不单凭品牌也不单凭名字。"),
        ("TXT", "②规格强一致(单件规格)", "关键：先比【单件规格base】而不是总量。3L稻米油 vs 250ml×12牛奶，总量都是3000ml但单件规格不同(3L vs 250ml)，判为规格不同、基本排除——避免总量碰巧相等的误判。规格一致+18%、不同−22%。"),
        ("TXT", "③单位文字差异归一", "g/克、kg/千克/公斤、ml/毫升/升、斤 全部归一（10斤=5000g、1kg=1000g）。写法不同不影响判定。"),
        ("TXT", "④数量差异=同款不同包装", "500ml*2、150ml*3、30瓶、40瓶 都识别为件数n；单件规格相同仅件数不同的（Spes 500ml*2 vs 京东500ml），视为同款不同包装、不判规格不同，标「核实包装」。ONE-TO-MANY：一个禄品对多个京东品时，按【规格相似度优先】排序（规格+件数全同 > 同规格不同件数 > 未知 > 规格不同）。"),
        ("TXT", "⑤双向价格反证", "总价按件数折成单件价，我方/京东单件价比 >2.5倍或<0.4倍扣15%、标「单价差大,疑非同款」。数量已换算，Spes双瓶vs单瓶不会误伤。"),
        ("TXT", "⑥品类硬约束(本次新增)", "从品名识别品类（食用油/乳制品/洗发/益生菌/米面杂粮…）。两边品类都识别到却不同（稻米油 vs 纯牛奶）→ 扣55%、标「品类不同,基本可排除」、整行标红。这条直接修掉了第10行的误判。"),
        ("TXT", "为什么不用严格同名", "严格同名只命中 2 个：京东名堆营销词、品牌别名与规格写法都不同。故用可解释的多因子打分，每项差异写成质疑点。"),
        ("TXT", "校准结果", f"用第一层已确认同款反算：中位 {(calib[len(calib)//2] if calib else 0):.0%}。真同款名字差异本就大，分数仅供排序，务必结合规格/件数/价格/品类核对判断。"),
        ("SEC", "四、统一筛选：先看什么、次看什么", ""),
        ("TXT", "复核优先级", "《02_待确认候选》第一列：①优先（品牌+型号+规格一致）→ ②次看（价差大/同规格不同件数/名称较像）→ ③存疑（品牌或规格存疑）→ ④排除（品类不同或规格+型号都不符）。已按此排序，①排最前。"),
        ("TXT", "重点看什么", "第二列：每条候选都写明您应优先核的点（核图片/核包装件数/核价格口径/直接排除），照着看即可，不用每条都从头判断。"),
        ("TXT", "人工确认", "在《02_待确认候选》的「人工确认」列填“是/否/疑似”，必要时在备注写理由。填完后告诉我，我会把确认=是的并入第一层，生成一份完整的《同品关系》主表。"),
        ("TXT", "只读声明", "全程只读源文件，未修改任何源表、未写库。确认结果只写在新工作簿，不回源。"),
        ("TXT", "生成时间", datetime.now().isoformat(timespec="seconds")),
    ]
    r = 3
    for kind, k, v in lines:
        if kind == "SEC":
            c = ws.cell(row=r, column=1, value=k)
            c.font = Font(bold=True, size=12, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=C_HEADER)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3); ws.row_dimensions[r].height = 24
        else:
            a = ws.cell(row=r, column=2, value=k); a.font = Font(bold=True, color=C_HEADER)
            a.alignment = Alignment(vertical="top", wrap_text=True)
            b = ws.cell(row=r, column=3, value=v); b.alignment = Alignment(vertical="top", wrap_text=True)
            b.font = Font(color="333333"); ws.row_dimensions[r].height = 32 if len(v) < 70 else 46
        r += 1
    ws.freeze_panes = "A3"

    # ---- 01 已确认同品 ----
    ws = wb.create_sheet("01_已确认同品")
    hdr = ["禄品P码", "禄品商品名称", "二级类目", "京东SKU", "京东商品名称", "京东本网P码",
           "禄品规格", "京东规格", "禄品件数", "京东件数", "单件价之比(我/京东)",
           "京东价来源", "京东市场售卖价(元)", "我们采用成本(元)", "我们输出价(元)",
           "我们价-京东价(元)", "价差比例", "京东价时间(VOP/历史)", "历史更新时间", "确认依据"]
    ws.append(hdr); style_header(ws, 1, [C_HEADER] * len(hdr), height=46)
    for x in l1:
        ws.append([x["p"], x["name"], x["cat"], x["sku"], x["jname"], x["vd_code"],
                   x["lp_spec"], x["vp_spec"], x["lp_cnt"], x["vp_cnt"], x["unit_ratio"],
                   x["jd_src"], x["jd_price"], x["cost"], x["our"], x["gap"], x["gap_pct"],
                   x["captured"], x["updated"], "历史表人工标同品同规格=是"])
    money = {13, 14, 15, 16}; pct = {17}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            c.border = BORDER; c.font = Font(size=10, color="333333")
            if c.column in money: c.number_format = "0.00"
            if c.column in pct: c.number_format = "0.0%"
        row[11].fill = PatternFill("solid", fgColor=C_CONF)
        # 单件价比值异常高亮（规则5）
        ur = row[10].value
        if isinstance(ur, (int, float)) and (ur > PRICE_RATIO_HI or ur < PRICE_RATIO_LO):
            row[10].fill = PatternFill("solid", fgColor=C_WARN)
    set_widths(ws, [14, 42, 14, 16, 42, 14, 12, 12, 10, 10, 16,
                    14, 18, 18, 16, 16, 12, 20, 20, 30])
    ws.freeze_panes = "C2"; ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    # ---- 02 待确认候选 ----（复核优先级 + 重点看什么 放在最前，统一筛选）
    ws = wb.create_sheet("02_待确认候选")
    hdr = ["复核优先级", "重点看什么", "禄品P码", "禄品商品名称", "二级类目", "候选序", "候选来源",
           "候选京东SKU", "候选京东商品名称", "京东品牌", "相似度", "相似度档",
           "品牌核对", "型号核对", "品类核对", "禄品规格", "京东规格", "禄品件数", "京东件数",
           "规格核对", "价格核对", "质疑点",
           "京东市场价(元)", "我们成本(元)", "我们输出价(元)", "京东本网P码", "价时间(VOP/历史)",
           "人工确认", "人工备注"]
    ws.append(hdr); style_header(ws, 1, [C_HEADER] * len(hdr), height=46)
    for x in l2:
        band = "高" if x["sim"] >= 0.85 else ("中" if x["sim"] >= 0.70 else "低")
        vtime = x["captured"] or x["hist_time"]
        ws.append([x["level"], x["focus"], x["p"], x["name"], x["cat"], x["rank"], x["source"],
                   x["sku"], x["jname"], x["brand"], x["sim"], band,
                   x["brand_flag"], x["model_flag"], x["cat_flag"], x["lp_spec"], x["vp_spec"],
                   x["lp_cnt"], x["vp_cnt"], x["spec_flag"], x["price_flag"], x["doubts"],
                   x["jd_price"], x["cost"], x["our"], x["bl_code"], vtime,
                   x["confirm"], x["note"]])
    money = {23, 24, 25}; pct = {11}
    lv_fill = {"①优先": C_HIGH, "②次看": C_MID, "③存疑": C_LOW, "④排除": C_WARN}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            c.border = BORDER; c.font = Font(size=10, color="333333")
            if c.column in money: c.number_format = "0.00"
            if c.column in pct: c.number_format = "0.0%"
        # 复核优先级高亮
        lv = str(row[0].value or "")
        if lv in lv_fill:
            row[0].fill = PatternFill("solid", fgColor=lv_fill[lv])
            row[0].font = Font(bold=True, size=10, color="333333")
        bc = row[11]
        bc.fill = PatternFill("solid", fgColor={"高": C_HIGH, "中": C_MID, "低": C_LOW}[bc.value])
        if "不一致" in str(row[12].value) or "不同" in str(row[13].value):
            row[21].fill = PatternFill("solid", fgColor=C_WARN)
        # 品类冲突整行标红
        if "冲突" in str(row[14].value or ""):
            for c in row[:15]:
                c.fill = PatternFill("solid", fgColor=C_WARN)
        # 规格不同/单价差大 高亮
        sf = str(row[19].value or ""); pf = str(row[20].value or "")
        if "不同" in sf:
            row[19].fill = PatternFill("solid", fgColor=C_WARN)
        if "差大" in pf:
            row[20].fill = PatternFill("solid", fgColor=C_WARN)
        row[27].fill = PatternFill("solid", fgColor=C_CAND)  # 人工确认列高亮
    set_widths(ws, [12, 46, 14, 38, 12, 8, 14, 15, 40, 18, 10, 10,
                    11, 22, 26, 12, 12, 10, 10, 24, 18, 40,
                    15, 14, 14, 14, 18, 12, 24])
    ws.freeze_panes = "D2"; ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    # ---- 03 字段与规则 ----
    ws = wb.create_sheet("03_字段与规则")
    set_widths(ws, [4, 24, 100])
    ws["A1"] = "字段含义与判定规则"; ws["A1"].font = Font(bold=True, size=14, color=C_HEADER)
    rows_txt = [
        ("禄品P码", "我们禄品体系自己的商品编码，定价基准。"),
        ("京东SKU / 京东本网P码", "京东商品编码 / 京东品在“本来体系”里对应的另一个本网 P 码——与禄品 P 码是两套。"),
        ("复核优先级", "统一筛选列：①优先=品牌+型号+规格一致；②次看=价差大/同规格不同件数/名称较像；③存疑=品牌或规格存疑；④排除=品类不同或规格+型号都不符。已按此排序，①排最前，绿→黄→橙→红。"),
        ("重点看什么", "每条候选写明您应优先核的点（核图片/核包装件数/核价格口径/直接排除），照着看即可，不用每条从头判断。"),
        ("京东市场售卖价", "京东 C 端客户看到的价格（jd_price）。第一层：VOP 实时价优先、缺失用历史表价；第二层：VOP 实时 jd_price。"),
        ("相似度", "按规则综合：字符序列(22%)+二字词(22%)+品牌(±24/25%)+型号(±14/8%)+规格强一致(±18/22%)+单价反证(−15%)+品类冲突(−55%)。归一到 0–100%，仅排序参考。"),
        ("品类核对", "从品名识别的品类（食用油/乳制品/洗发/益生菌/米面杂粮…）。两边都识别到却不同→标「品类冲突」、整行标红、基本排除。修掉稻米油vs牛奶这类误判。"),
        ("禄品/京东规格", "单件规格base（不是总量！），已单位归一（g/克、kg、ml/毫升、斤、1kg=1000g）。先比单件规格：3L vs 250ml 即使总量都3000ml也算规格不同。"),
        ("禄品/京东件数", "从 *2、*3、2瓶装、双支、30瓶、500ml+500ml 等识别的包装件数。单件规格相同、仅件数不同的视为同款不同包装（Spes 500ml*2 vs 500ml）。"),
        ("规格核对", "规格一致 / 同规格不同件数 / 规格不同 / 无规格。注意比的是单件规格，规格不同会重扣分。"),
        ("价格核对", "我方与京东「单件价之比」（已按件数折算）。>2.5倍或<0.4倍标「单价差大」，价格差太大则降低相似度——反过来说明可能不是同一个品。"),
        ("相似度档", "高≥85% / 中 70–85% / 低 40–70%。连真同款的相似度中位也不高（京东名堆营销词），低档不等于不是同款，务必结合规格/件数/价格核对与质疑点判断。门槛40%只为尽量不漏。"),
        ("质疑点", "系统发现的差异：品牌不一致、型号不同、规格不同/同规格不同件数、单价差大、名称相似度偏低。帮您快速判断。"),
        ("人工确认", "请您填“是/否/疑似”。只有填“是”的，我才会并入完整同品关系主表。绝不自动判定。"),
        ("我们输出价-京东价", "正=我们比京东贵，负=我们比京东便宜。用于看定价相对京东市场的位置。"),
    ]
    r = 3
    for k, v in rows_txt:
        ws.cell(row=r, column=2, value=k).font = Font(bold=True, color=C_HEADER)
        ws.cell(row=r, column=3, value=v).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=3).font = Font(color="333333")
        ws.row_dimensions[r].height = 30
        r += 1
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=3):
        for c in row: c.border = BORDER
    wb.save(OUTPUT)


if __name__ == "__main__":
    main()
