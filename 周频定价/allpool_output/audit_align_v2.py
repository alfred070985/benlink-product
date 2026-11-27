#!/usr/bin/env python3
"""v2 精排对齐结果逐行审查（只读分析，不改算法、不改货盘）。

四条审查线：
  A. 猫超 69 码真条码独立验证（客观真值）——真同款召回/被压分情况
  B. 全量采信对一致性扫描：品类冲突、品牌冲突补漏、规格矛盾、数字冲突
  C. 价格异常：折算价严重倒挂、精排分与向量分背离
  D. 分层抽样明细导出，供人工逐行判读

输出：/tmp/audit_align_v2.xlsx（初稿，供人工判读后再决定是否入库）
"""
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "pricing_output"))

from latest import latest  # noqa: E402
from build_compare_board import (  # noqa: E402
    load_names, load_bl_cost, load_prices, cnt_from_attr,
)

ACCEPT = 0.8
RESCUE_FLOOR = 0.75
OUT = Path("/tmp/audit_align_v2.xlsx")
EXCL_F = ROOT / "交付物/2026-W48/03_货盘底表/5_跨源对齐/分销池_ERP京东品.jsonl"

# ---------- 品类词典：词 -> 大类（同义词归并，避免误报） ----------
CAT = {}
def _cat(group, words):
    for w in words:
        CAT[w] = group

_cat("米", ["大米", "稻米", "香米", "丝苗米", "珍珠米", "粳米", "籼米", "糯米", "糙米", "小米", "杂粮", "米饭", "五常米", "东北米"])
_cat("油", ["食用油", "稻米油", "葵花籽油", "调和油", "花生油", "菜籽油", "橄榄油", "玉米油", "大豆油", "色拉油", "米糠油", "芝麻油", "亚麻籽油"])
_cat("奶", ["牛奶", "纯牛奶", "酸奶", "奶粉", "奶酪", "乳饮料", "调制乳", "鲜奶", "奶片"])
_cat("洗护", ["洗发水", "洗发露", "护发素", "沐浴露", "洗洁精", "洗衣液", "牙膏", "肥皂", "香皂", "洗手液", "洗面奶"])
_cat("纸品", ["纸巾", "卫生纸", "抽纸", "湿巾", "卷纸", "厨房纸", "面巾纸"])
_cat("水果", ["奇异果", "猕猴桃", "苹果", "香蕉", "橙子", "橘子", "梨", "葡萄", "草莓", "西瓜", "榴莲", "蓝莓", "车厘子", "芒果", "柚子", "桃", "柠檬", "火龙果"])
_cat("肉蛋", ["猪肉", "牛肉", "羊肉", "鸡肉", "鸡蛋", "鸭蛋", "火腿", "培根", "鸭肉", "鹅肉", "肉丸", "香肠"])
_cat("水产", ["虾", "鱼", "海鲜", "蟹", "贝", "鱿鱼", "海参", "带鱼", "三文鱼", "生蚝", "扇贝"])
_cat("饮料", ["咖啡", "果汁", "饮料", "矿泉水", "啤酒", "白酒", "红酒", "葡萄酒", "可乐", "茶饮", "纯净水", "苏打水"])
_cat("茶", ["茶叶", "绿茶", "红茶", "乌龙茶", "普洱茶", "花茶", "黑茶", "白茶"])
_cat("粮面", ["面粉", "面条", "挂面", "麦片", "燕麦", "方便面", "米粉", "意面", "荞麦", "玉米糁"])
_cat("干货", ["木耳", "香菇", "红枣", "枸杞", "坚果", "花生", "瓜子", "干果", "桂圆", "莲子", "银耳", "腐竹", "粉丝"])
_cat("蔬菜", ["白菜", "萝卜", "土豆", "番茄", "西红柿", "黄瓜", "青菜", "韭菜", "西兰花", "胡萝卜", "洋葱", "大蒜", "生姜"])
_cat("烘焙", ["面包", "蛋糕", "饼干", "糕点", "月饼", "蛋挞", "吐司", "点心"])
_cat("家电日用", ["电池", "军刀", "水果刀", "充电", "数据线", "插座", "灯泡", "收纳", "剪刀", "锅", "水杯", "保温杯", "雨伞"])
_cat("保健", ["维生素", "蛋白粉", "鱼油", "钙片", "保健品", "益生菌", "胶原"])
_cat("宠物", ["猫粮", "狗粮", "宠物", "猫砂", "狗零食"])
_cat("母婴", ["尿裤", "纸尿裤", "奶粉", "辅食", "奶瓶", "婴儿"])


def cats_of(name):
    """品类词抽取。校准：口味/香型语境（如"葡萄味""樱花+白桃"）不算品类，
    否则会把"果蔬汁（葡萄味）vs 果汁"误报成品类冲突。"""
    s = str(name or "")
    out = set()
    for w, g in CAT.items():
        if w not in s:
            continue
        for m in re.finditer(re.escape(w), s):
            tail = s[m.end(): m.end() + 1]
            if tail and tail[0] in "味香型":      # 葡萄味 / 白桃香
                continue
            if g in ("水果", "蔬菜") and re.search(
                    r"(汁|饮料|酱|香|味|沐浴|洗发|洗衣|皂|茶|酒|口红|牙膏)", s):
                continue                            # 香型/成分，非真实品类
            out.add(g)
            break
    return out


def brand_head(name):
    """取品名开头 2-4 字作为品牌锚点。校准：首字相同视为同一品牌族，避免
    "阳澄湖 vs 阳澄江南"这类近似品牌误报。"""
    s = re.sub(r"^【[^】]*】", "", str(name or "")).strip()
    m = re.match(r"^([\u4e00-\u9fa5A-Za-z]{2,4})", s)
    return m.group(1) if m else ""


def num_tokens(name):
    """抽取品名中的规格数字+单位（如 5kg / 500ml / 16粒），用于数字冲突检测。"""
    s = str(name or "").lower()
    return set(re.findall(r"(\d+(?:\.\d+)?)\s*(kg|g|ml|l|粒|袋|桶|盒|支|片|包|斤)", s))


def load_align(fn):
    rows = []
    with (HERE / fn).open(encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            rows.append(json.loads(line))
    return rows


def main():
    names = load_names()
    costs = load_bl_cost()
    excl = {str(json.loads(l)["pbasic"]) for l in EXCL_F.open(encoding="utf-8") if l.strip()}
    prices = {w: load_prices(w) for w in ("jd", "mc")}
    align = {w: load_align(f"对齐_v2_bl_{w}.jsonl") for w in ("jd", "mc")}

    # ---------- 全量逐行扫描 ----------
    detail = []          # 所有采信行 + 可疑标记
    sig_count = Counter()
    for w in ("jd", "mc"):
        pr = prices[w]
        for o in align[w]:
            a, b = str(o["a"]), str(o["b"])
            s = float(o["sim"])
            if a in excl:
                continue
            tier = "采信" if s >= ACCEPT else "边界"
            an = o.get("a_name") or names.get(a, "")
            bn = o.get("b_name") or pr.get(b, {}).get("name", "")
            aa, bb = o.get("a_attr") or {}, o.get("b_attr") or {}
            cb = pr.get(b, {}).get("channel") or pr.get(b, {}).get("hand")
            cnt_a = cnt_from_attr(aa, an)
            cnt_b = cnt_from_attr(bb, bn)
            adj = round(cb / cnt_b * cnt_a, 2) if cb and cnt_b and cnt_a else None
            cost = costs.get(a)
            sigs = []
            # 品类冲突信号（2026-11-27）：复用与硬约束同一套强品类词判定，
            # 口径一致且避免旧 cats_of 词典"金龙鱼→水产/葡萄味→水果"的误报。
            from category_classify import cat_conflict as _cc, is_combo as _ico
            if _cc(an, bn, _ico(an) or _ico(bn)):
                sigs.append("品类冲突")
            bh_a, bh_b = brand_head(an), brand_head(bn)
            if (bh_a and bh_b and bh_a != bh_b and bh_a[0] != bh_b[0]
                    and not (bh_a in bn or bh_b in an)):
                sigs.append("品牌词冲突")
            if aa.get("base") and bb.get("base") and aa["base"] != bb["base"]:
                sigs.append("净含量不等")
            if aa.get("pack_explicit") and bb.get("pack_explicit") and aa.get("count") != bb.get("count"):
                sigs.append("件数不等")
            if o.get("spec_penalty", 0) > 0:
                sigs.append(f"规格矛盾减分{int(o['spec_penalty']*100)}")
            na, nb = num_tokens(an), num_tokens(bn)
            if na and nb and not (na & nb):
                sigs.append("规格数字冲突")
            if adj and cost and cost > 0 and adj / cost < 0.5:
                sigs.append("折算价严重倒挂")
            if s >= ACCEPT and float(o.get("sim_vec", 0)) < 0.6:
                sigs.append("精排高但向量低")
            if tier == "采信" or sigs:
                for x in sigs:
                    sig_count[(tier, x)] += 1
                detail.append({
                    "货源": "京东" if w == "jd" else "猫超", "pbasic": a, "对方编号": b,
                    "本来品名": an, "对照品名": bn,
                    "成本": cost, "渠道价": cb, "本来件数": cnt_a, "对方件数": cnt_b,
                    "折算价": adj, "精排分": s, "向量分": float(o.get("sim_vec", 0)),
                    "档位": tier, "候选序": o.get("rank"), "信号": "、".join(sigs),
                })

    # ---------- 真值验证（猫超 69 码） ----------
    truth_rows = []
    gt = [json.loads(l) for l in (HERE / "对齐_猫超69码.jsonl").open(encoding="utf-8") if l.strip()]
    mcbar = latest("supply_pull/output/猫超商品宽表SKU级_*.csv")
    bar2pid = {}
    with open(mcbar, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            bc = (r.get("sku_barcode") or "").strip()
            pid = str(r.get("product_id") or "")
            if bc and pid and bc not in bar2pid:
                bar2pid[bc] = pid
    # 本网 pbasic -> psys -> name
    pool = latest("pricing_output/标准池_四站全上架有库存_*.jsonl")
    info = latest("pricing_output/源数据_标准池商品信息_*.jsonl")
    p2b = {str(json.loads(l)["productSysNo"]): str(json.loads(l)["productBasicSysNo"])
           for l in Path(pool).open(encoding="utf-8") if l.strip()}
    b2p = {v: k for k, v in p2b.items()}
    nm = {}
    for l in Path(info).open(encoding="utf-8"):
        if l.strip():
            o = json.loads(l)
            nm[str(o["productSysNo"])] = o.get("name", "")
    mc_by_pb = defaultdict(list)
    for o in align["mc"]:
        mc_by_pb[str(o["a"])].append(o)
    stat = Counter()
    for g in gt:
        pb = str(g["pbasic"])
        pid = bar2pid.get(str(g["barcode"]))
        if not pid or pb in excl:
            continue
        psys = b2p.get(pb)
        if not psys:
            continue
        blname = nm.get(psys, "")
        cands = mc_by_pb.get(pb, [])
        hit = [o for o in cands if str(o["b"]) == pid]
        stat["真同款可比对"] += 1
        if hit:
            o = hit[0]
            stat["真同款被召回"] += 1
            t = "采信" if o["sim"] >= ACCEPT else "边界"
            stat[f"真同款-{t}"] += 1
            if t == "边界":
                truth_rows.append({
                    "pbasic": pb, "本网品名": blname, "真猫超编号": pid,
                    "真猫超品名": o.get("b_name", ""), "精排分": o["sim"],
                    "向量分": o.get("sim_vec"), "规格减分": o.get("spec_penalty", 0),
                    "判定": "真同款被压到边界", "信号": "、".join(sorted(cats_of(blname))) + " / " + "、".join(sorted(cats_of(o.get("b_name", "")))),
                })
        else:
            stat["真同款未召回"] += 1

    # ---------- 写工作簿 ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "00_审查摘要"
    ws.append(["审查项", "值"])
    ws.append(["全量候选行(采信+有信号)", len(detail)])
    ws.append(["采信行数", sum(1 for d in detail if d["档位"] == "采信")])
    ws.append(["边界且有信号行数", sum(1 for d in detail if d["档位"] == "边界")])
    for (tier, sig), n in sorted(sig_count.items(), key=lambda x: -x[1]):
        ws.append([f"{tier}·{sig}", n])
    ws.append(["— 69码真值验证 —", ""])
    for k, v in stat.items():
        ws.append([k, v])

    ws = wb.create_sheet("01_可疑采信行")
    hdr = ["货源", "pbasic", "对方编号", "本来品名", "对照品名", "成本", "渠道价",
           "本来件数", "对方件数", "折算价", "精排分", "向量分", "档位", "候选序", "信号"]
    ws.append(hdr)
    sus = [d for d in detail if d["档位"] == "采信" and d["信号"]]
    sus.sort(key=lambda d: (("品类冲突" in d["信号"]), ("折算价严重倒挂" in d["信号"]), d["精排分"]), reverse=True)
    for d in sus[:5000]:
        ws.append([d[h] for h in hdr])

    ws = wb.create_sheet("02_真同款被压边界")
    hdr2 = ["pbasic", "本网品名", "真猫超编号", "真猫超品名", "精排分", "向量分", "规格减分", "判定", "信号"]
    ws.append(hdr2)
    for r in truth_rows[:3000]:
        ws.append([r[h] for h in hdr2])

    ws = wb.create_sheet("03_全部采信行")
    hdr3 = ["货源", "pbasic", "对方编号", "本来品名", "对照品名", "成本", "渠道价",
            "本来件数", "对方件数", "折算价", "精排分", "向量分", "档位", "候选序", "信号"]
    ws.append(hdr3)
    for d in detail:
        if d["档位"] == "采信":
            ws.append([d[h] for h in hdr3])

    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        for c in sheet[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="305496")
    wb.save(OUT)
    print(f"已生成 {OUT}")
    print("摘要:", dict(stat))
    print("信号计数:", dict(sig_count))
    print(f"可疑采信 {len(sus)} 行, 真同款被压边界 {len(truth_rows)} 行")


if __name__ == "__main__":
    main()
