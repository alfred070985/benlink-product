#!/usr/bin/env python3
"""在《禄品分销商品定价_含新京东价》基础上，新增“选品结论 + 建议加价档 + 统计”。

规则（用户2026-11-16拍板）：
  - 不做直客，须给下游渠道留空间；空间率 =(京东价-输出价)/京东价
  - 选品：有京东价且最低档(档①)空间率>=10% -> 选中；<10%或倒挂 -> 不选；无成本 -> 不选
          无京东价但有成本 -> 参与(无对标价)，照常给档位
  - 选档：在选中/参与品里，从最高档往下找第一个满足“空间率>=10%”的档；
          无京东价(无法算空间) -> 取最高档(纯成本加成)，备注说明
  - 空间下限 10%（用户采纳“你的建议”）

另存新文件，不改动源表。路径由 __file__ 推导。
"""
import os
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

HERE = Path(__file__).resolve().parent
PRICING_DIR = Path(os.environ.get("PRICING_DIR", HERE))
TODAY = datetime.now().strftime("%Y-%m-%d")
SRC = PRICING_DIR / f"禄品分销商品定价_含新京东价_{TODAY}.xlsx"
OUT = PRICING_DIR / f"禄品分销商品定价_含选品建议_{TODAY}.xlsx"

SPACE_FLOOR = 0.10

C_HEADER = "1F4E79"
C_WHITE = "FFFFFF"
C_PICK = "E2EFDA"     # 选中-绿
C_NOPICK = "FCE4D6"   # 不选-橙
C_SUGGEST = "FFF2CC"  # 建议档高亮-黄
C_PENDING = "F2F2F2"


def space(jd, out):
    if jd in (None, "") or out in (None, ""):
        return None
    jd = float(jd)
    return (jd - float(out)) / jd if jd > 0 else None


def decide(cost, jd, tiers):
    """tiers=[(ratio,price,gross,gm)*3] 返回 dict。"""
    ratios = [t[0] for t in tiers]
    prices = [t[1] for t in tiers]
    spaces = [space(jd, p) for p in prices]
    has_cost = cost not in (None, "")
    has_jd = jd not in (None, "")

    if not has_cost:
        return dict(sel="不选", reason="无成本·无法出价", tier=None, ratio=None,
                    out=None, s_tier=None, spaces=spaces, note="成本缺失"), spaces

    if not has_jd:
        # 无对标价：照常参与，取最高档，纯成本加成
        idx = 2
        return dict(sel="参与(无对标价)", reason="无京东价·按成本加成", tier=idx + 1,
                    ratio=ratios[idx], out=prices[idx], s_tier=None, spaces=spaces,
                    note="无对标价·空间待补"), spaces

    # 有价：选品看最低档空间
    s1 = spaces[0]
    if s1 is None or s1 < SPACE_FLOOR:
        tag = "倒挂" if (s1 is not None and s1 < 0) else "空间不足"
        return dict(sel="不选", reason=f"档①空间{s1:.0%}<10%" if s1 is not None else "无空间",
                    tier=None, ratio=None, out=None, s_tier=None, spaces=spaces, note=tag), spaces

    # 选档：从最高档往下找第一个空间>=10%
    chosen = None
    for i in range(2, -1, -1):
        if spaces[i] is not None and spaces[i] >= SPACE_FLOOR:
            chosen = i
            break
    if chosen is None:
        chosen = 0
    return dict(sel="选中", reason=f"档①空间{s1:.0%}≥10%", tier=chosen + 1,
                ratio=ratios[chosen], out=prices[chosen], s_tier=spaces[chosen],
                spaces=spaces, note=""), spaces


def main():
    print(f"[{TODAY}] 新增选品结论与建议加价档", flush=True)
    wb = openpyxl.load_workbook(SRC)
    ws = wb["04_定价结果"]

    # 现有：第1行分组、第2行表头、第3行起数据；A..U 共21列
    group_row, hdr_row = 1, 2
    new_cols = ["选品结论", "判断依据", "建议档位", "建议加价比例", "建议输出价(元)",
                "建议档空间率", "档①空间率", "档②空间率", "档③空间率", "备注"]
    start = 22
    ws.cell(row=group_row, column=start, value="选品与建议（不做直客·留渠道空间）")
    for c, name in enumerate(new_cols, start):
        cell = ws.cell(row=hdr_row, column=c, value=name)
        cell.fill = PatternFill("solid", fgColor=C_HEADER)
        cell.font = Font(color=C_WHITE, bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.cell(row=group_row, column=start).font = Font(bold=True, color=C_HEADER)

    sel_stat = Counter()
    tier_stat = Counter()
    ship_tier = Counter()
    results = []

    for r in range(hdr_row + 1, ws.max_row + 1):
        pc = ws.cell(row=r, column=1).value
        if not pc:
            continue
        ship = ws.cell(row=r, column=4).value
        cost = ws.cell(row=r, column=5).value
        jd = ws.cell(row=r, column=6).value
        tiers = [
            (ws.cell(row=r, column=9).value, ws.cell(row=r, column=10).value),
            (ws.cell(row=r, column=13).value, ws.cell(row=r, column=14).value),
            (ws.cell(row=r, column=17).value, ws.cell(row=r, column=18).value),
        ]
        tiers_full = [(rt, pr, None, None) for rt, pr in tiers]
        dec, spaces = decide(cost, jd, tiers_full)
        vals = [dec["sel"], dec["reason"],
                f"档{dec['tier']}" if dec["tier"] else "—",
                dec["ratio"], dec["out"], dec["s_tier"],
                spaces[0], spaces[1], spaces[2], dec["note"]]
        for c, v in enumerate(vals, start):
            cell = ws.cell(row=r, column=c, value=v)
            if c in (25, 27, 28, 29, 30):
                cell.number_format = "0.00%"
            if c == 26:
                cell.number_format = "0.00"
        # 高亮选品结论
        sc = ws.cell(row=r, column=start)
        if dec["sel"] == "选中":
            sc.fill = PatternFill("solid", fgColor=C_PICK)
        elif dec["sel"] == "不选":
            sc.fill = PatternFill("solid", fgColor=C_NOPICK)
        else:
            sc.fill = PatternFill("solid", fgColor=C_PENDING)
        # 高亮建议的那一档（比例/输出价/毛利/毛利率 4列）
        if dec["tier"]:
            base = 8 + 4 * dec["tier"]  # 档1->12? 用偏移映射
            base = {1: 9, 2: 13, 3: 17}[dec["tier"]]
            for cc in range(base, base + 4):
                ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=C_SUGGEST)
        sel_stat[dec["sel"]] += 1
        if dec["tier"]:
            tier_stat[f"档{dec['tier']}"] += 1
            ship_tier[(ship, f"档{dec['tier']}")] += 1
        results.append((ship, dec, spaces))

    # 列宽
    widths = [14, 20, 10, 14, 16, 14, 12, 12, 12, 18]
    for i, w in enumerate(widths, start):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A2:{openpyxl.utils.get_column_letter(start+len(new_cols)-1)}{ws.max_row}"

    # ---- 补充 00 说明页 ----
    ws0 = wb["00_怎么读这张表"]
    for r in [
        ["本版新增：选品与建议档", "04表最右侧 V列起为新增的“选品与建议”区", ""],
        ["选品结论", "选中／不选／参与(无对标价)。不做直客,须给渠道留空间,空间不足或倒挂的不选", ""],
        ["建议档位", "在选中品里,取仍能留≥10%空间的最高加价档;该档比例/输出价已黄色高亮", ""],
    ]:
        ws0.append(r)

    # ---- 补充 01 规则页 ----
    ws1 = wb["01_定价规则"]
    for r in [
        ["SEC", "六、选品与建议档（本版新增，用户2026-11-16拍板）", ""],
        ["前提", "不做直客", "定价后要留给下游渠道利润空间，空间率=(京东价-输出价)/京东价"],
        ["选品", "空间率≥10%才选", "有京东价：最低档(档①)空间≥10%→选中；<10%或倒挂→不选；无成本→不选；无京东价→参与(无对标价)"],
        ["选档", "空间最大化可取的最高档", "选中品从最高档往下找第一个仍满足空间≥10%的档；无京东价无法算空间则取最高档(纯成本加成)"],
        ["下限", "10%", "渠道空间下限10%（用户采纳建议值）；表中“建议档空间率”“档①②③空间率”可逐行核对"],
    ]:
        ws1.append(r)
        if r[0] == "SEC":
            for c in ws1[ws1.max_row]:
                c.fill = PatternFill("solid", fgColor="D9E1F2")
                c.font = Font(bold=True, color=C_HEADER)

    # ---- 新增统计页 ----
    st = wb.create_sheet("07_选品与档位统计")
    st.append(["选品与建议加价档 统计（不做直客·须留渠道空间）", "", "", ""])
    st.cell(row=1, column=1).font = Font(bold=True, size=13, color=C_HEADER)
    st.append([])
    st.append(["规则", "空间下限10%：最低档空间<10%或倒挂→不选；选中品从最高档往下取空间仍≥10%的最高档", "", ""])
    st.append([])
    st.append(["一、选品结论分布", "数量", "占比", ""])
    n = len(results)
    for k in ["选中", "参与(无对标价)", "不选"]:
        v = sel_stat.get(k, 0)
        st.append([k, v, v / n if n else 0, ""])
    st.append(["合计", n, 1, ""])
    st.append([])
    st.append(["二、建议档位分布（选中+参与）", "数量", "占比", ""])
    tot_tier = sum(tier_stat.values())
    for k in ["档1", "档2", "档3"]:
        v = tier_stat.get(k, 0)
        st.append([k, v, v / tot_tier if tot_tier else 0, ""])
    st.append([])
    st.append(["三、发货类型 × 建议档位", "档1", "档2", "档3"])
    for ship in ["直发", "仓发"]:
        st.append([ship, ship_tier.get((ship, "档1"), 0),
                   ship_tier.get((ship, "档2"), 0), ship_tier.get((ship, "档3"), 0)])
    st.append([])
    st.append(["四、选中品的建议档空间率分布", "数量", "", ""])
    sb = Counter()
    for ship, dec, spaces in results:
        if dec["sel"] == "选中" and dec["s_tier"] is not None:
            s = dec["s_tier"]
            k = "≥30%" if s >= .3 else "20~30%" if s >= .2 else "10~20%" if s >= .1 else "<10%"
            sb[k] += 1
    for k in ["≥30%", "20~30%", "10~20%", "<10%"]:
        st.append([k, sb.get(k, 0), "", ""])
    st.append([])
    st.append(["口径说明", "空间率=(京东价-输出价)/京东价；京东价为CSV售价(不同规折算后)；成本与加价档未改动", "", ""])
    for row in st.iter_rows(min_row=5):
        if isinstance(row[2].value, float):
            row[2].number_format = "0.0%"
    st.column_dimensions["A"].width = 40
    st.column_dimensions["B"].width = 60
    st.column_dimensions["C"].width = 14
    st.column_dimensions["D"].width = 14

    wb.save(OUT)
    print(f"产出: {OUT.name}", flush=True)
    print("选品:", dict(sel_stat), flush=True)
    print("建议档:", dict(tier_stat), flush=True)
    print("发货×档:", dict(ship_tier), flush=True)


if __name__ == "__main__":
    main()
