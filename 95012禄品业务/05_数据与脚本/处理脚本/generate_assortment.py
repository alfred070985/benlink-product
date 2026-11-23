#!/usr/bin/env python3
import json
import zipfile
from collections import Counter
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
OUTPUT = ROOT / "来福商品池_一件代发收益及入盘建议_2026-10-30.xlsx"
LOW_FULFILLMENT = 3.0
NORMAL_FULFILLMENT = 5.0
TARGET_PROFIT = 3.0
MARKUP = 0.02
NEW_HEADERS = [
    "一件代发履约成本(元/单)", "一件代发加2%售价(元)", "加2%单均收益(元)",
    "现池售价单均收益(元)", "保本售价(元)", "目标售价(元)", "现售价与目标价差(元)",
    "入盘建议", "建议原因", "5元履约情景收益(元)", "5元情景建议",
]


def safe(value):
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def repair_double_escaped_xml(path):
    temp = path.with_suffix(".tmp.xlsx")
    replacements = 0
    with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(temp, "w") as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename.endswith(".xml"):
                count = data.count(b"_x005F_x005F_")
                if count:
                    replacements += count
                    data = data.replace(b"_x005F_x005F_", b"_x005F_")
            dst.writestr(info, data)
    temp.replace(path)
    return replacements


def round2(value):
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def advice(current, cost, fulfillment, target=TARGET_PROFIT):
    if current is None or cost is None or current <= 0 or cost <= 0:
        return "不建议入盘", None
    profit = current - cost - fulfillment
    if profit >= target:
        return "建议入盘", profit
    if profit > 0:
        return "谨慎入盘", profit
    return "不建议入盘", profit


def main():
    formula_wb = load_workbook(INPUT, read_only=True, data_only=False)
    value_wb = load_workbook(INPUT, read_only=True, data_only=True)
    main_f = formula_wb["二次定价"]
    main_v = value_wb["二次定价"]
    f_rows = main_f.iter_rows(values_only=True)
    v_rows = main_v.iter_rows(values_only=True)
    base_headers = list(next(f_rows))
    next(v_rows)
    headers = base_headers + NEW_HEADERS
    col = {name: i + 1 for i, name in enumerate(headers)}
    max_row = main_f.max_row

    workbook = xlsxwriter.Workbook(OUTPUT, {"constant_memory": True})
    workbook.set_calc_mode("auto")
    header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"})
    title_fmt = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E78"})
    section_fmt = workbook.add_format({"bold": True, "font_color": "#1F4E78"})
    input_fmt = workbook.add_format({"font_color": "#0000FF"})
    money_fmt = workbook.add_format({"num_format": "0.00"})
    pct_fmt = workbook.add_format({"num_format": "0.00%"})
    int_fmt = workbook.add_format({"num_format": "#,##0"})
    note_fmt = workbook.add_format({"bg_color": "#FFF2CC", "text_wrap": True, "valign": "top"})
    green_fmt = workbook.add_format({"bg_color": "#E2F0D9", "font_color": "#006100"})
    yellow_fmt = workbook.add_format({"bg_color": "#FFF2CC", "font_color": "#9C6500"})
    red_fmt = workbook.add_format({"bg_color": "#FCE4D6", "font_color": "#9C0006"})

    summary = workbook.add_worksheet("汇总")
    summary.hide_gridlines(2)
    summary.set_column("A:A", 34)
    summary.set_column("B:B", 20)
    summary.set_column("D:D", 30)
    summary.set_column("E:E", 95)
    summary.write(0, 0, "来福商品池一件代发收益及入盘建议", title_fmt)
    summary.write(1, 0, "核心结论", section_fmt)
    summary.write(1, 1, "入盘看每单绝对收益，不只看成本加价率；成本加2%与现池售价两种口径分别测算。")
    summary.write(2, 0, "低位情景", section_fmt)
    summary.write(2, 1, "履约3元/单，目标单均收益3元。")
    summary.write(3, 0, "压力情景", section_fmt)
    summary.write(3, 1, "履约5元/单，目标单均收益仍为3元。")
    summary.write(4, 0, "生成时间", section_fmt)
    summary.write(4, 1, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    summary.write(6, 0, "结果指标", section_fmt)
    summary.write(6, 1, "商品数", section_fmt)

    params = workbook.add_worksheet("入盘参数")
    params.hide_gridlines(2)
    params.set_column("A:A", 38)
    params.set_column("B:B", 20)
    params.set_column("D:D", 110)
    params.write(0, 0, "一件代发入盘参数（蓝色为可调整输入）", title_fmt)
    parameter_rows = [
        ("低位履约成本(元/单)", LOW_FULFILLMENT, "公开云仓报价低位区间的测算假设，不代表公司实际合同价。"),
        ("目标单均收益(元)", TARGET_PROFIT, "现池售价－采用成本－履约成本达到该值，才标记为“建议入盘”。"),
        ("一件代发成本加价率", MARKUP, "全部商品统一按一件代发加2%测算，不再使用原仓发重量分档。"),
        ("谨慎情景履约成本(元/单)", NORMAL_FULFILLMENT, "用于检验真实履约成本上升时的收益和建议。"),
    ]
    for r, (label, value, note) in enumerate(parameter_rows, start=1):
        params.write(r, 0, label)
        params.write_number(r, 1, value, pct_fmt if "加价率" in label else money_fmt)
        params.write(r, 3, note)
    notes = [
        "入盘建议以现池售价为可实现售价：单均收益≥3元为建议入盘；0＜收益＜3元为谨慎入盘；收益≤0或成本/售价缺失为不建议入盘。",
        "成本加2%单均收益单独列示。若现售价可入盘但加2%亏损，建议原因会提示“不能按成本加2%销售”。",
        "3元/单为低位履约假设，公开云仓服务报价通常由仓储、操作、包材、快递和增值服务构成，实际必须用合同价替换。",
        "目标售价＝采用成本＋低位履约成本＋目标单均收益；保本售价＝采用成本＋低位履约成本。",
        "本建议只判断单件经济收益，不替代品质、供应商稳定性、退货率、合规和销量审查。",
    ]
    for i, note in enumerate(notes, start=7):
        params.write(i, 0, note, note_fmt)

    # Keep the original cost-category parameters because the original formulas reference this sheet.
    for sheet_name in ("定价参数", "数据来源"):
        src_f = formula_wb[sheet_name]
        src_v = value_wb[sheet_name]
        dst = workbook.add_worksheet(sheet_name)
        dst.hide_gridlines(2)
        dst.set_column("A:A", 40)
        dst.set_column("B:B", 105)
        dst.set_column("C:C", 85)
        dst.set_column("D:E", 32)
        vf_rows = src_f.iter_rows(values_only=True)
        vv_rows = src_v.iter_rows(values_only=True)
        for r, (fr, vr) in enumerate(zip(vf_rows, vv_rows)):
            for c, value in enumerate(fr):
                if value is None:
                    continue
                if isinstance(value, str) and value.startswith("="):
                    dst.write_formula(r, c, value, None, vr[c] if vr[c] is not None else "")
                else:
                    dst.write(r, c, safe(value))

    # Add industry references to the source sheet after the copied rows.
    sources = workbook.get_worksheet_by_name("数据来源")
    source_start = formula_wb["数据来源"].max_row + 2
    sources.write(source_start, 0, "交通运输部／国家邮政局：2025年邮政行业运行情况", section_fmt)
    sources.write(source_start, 1, "https://www.mot.gov.cn/shuju/tongjishuju/youzheng/202601/t20260130_4199355.html")
    sources.write(source_start, 2, "2025年快递业务收入14939.3亿元、业务量1989.5亿件；全国平均单票收入约7.51元，仅作行业参照，不等同电商云仓合同价。")
    sources.write(source_start + 1, 0, "云仓服务商公开收费说明（低位假设参考）", section_fmt)
    sources.write(source_start + 1, 1, "http://www.sdyuncang.com/?industry/592.html")
    sources.write(source_start + 1, 2, "公开报价称整体约2–10元/单、3–6元/单居多；本表取3元作为低位情景，并另列5元压力情景。")

    ws = workbook.add_worksheet("二次定价")
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max_row - 1, len(headers) - 1)
    ws.set_row(0, 32)
    for c, name in enumerate(headers):
        ws.write(0, c, name, header_fmt)
    widths = {
        "商品编号": 16, "商品名称": 42, "商品大类": 16, "商品中类": 18, "商品小类": 20,
        "商品池售价": 14, "是否产地直发": 16, "成本取值口径": 28, "采用成本(元)": 16,
        "二次定价(元)": 16, "定价状态": 18, "一件代发履约成本(元/单)": 24,
        "一件代发加2%售价(元)": 23, "加2%单均收益(元)": 20, "现池售价单均收益(元)": 22,
        "保本售价(元)": 16, "目标售价(元)": 16, "现售价与目标价差(元)": 22,
        "入盘建议": 16, "建议原因": 58, "5元履约情景收益(元)": 22, "5元情景建议": 16,
    }
    for name, idx in col.items():
        ws.set_column(idx - 1, idx - 1, widths.get(name, 15))

    stats = Counter()
    for excel_row, (fr, vr) in enumerate(zip(f_rows, v_rows), start=2):
        out_row = excel_row - 1
        for c, value in enumerate(fr):
            cached = vr[c]
            if isinstance(value, str) and value.startswith("="):
                ws.write_formula(out_row, c, value, None, "" if cached is None else cached)
            elif value is not None:
                ws.write(out_row, c, safe(value))
        cost = vr[col["采用成本(元)"] - 1]
        current = vr[col["商品池售价"] - 1]
        cost = float(cost) if cost is not None else None
        current = float(current) if current is not None else None
        second_price = round2(cost * (1 + MARKUP)) if cost is not None and cost > 0 else None
        profit_second = second_price - cost - LOW_FULFILLMENT if second_price is not None else None
        decision, profit_current = advice(current, cost, LOW_FULFILLMENT)
        decision5, profit5 = advice(current, cost, NORMAL_FULFILLMENT)
        break_even = cost + LOW_FULFILLMENT if cost is not None and cost > 0 else None
        target_price = cost + LOW_FULFILLMENT + TARGET_PROFIT if break_even is not None else None
        target_diff = current - target_price if current is not None and target_price is not None else None
        if cost is None or cost <= 0:
            reason = "采用成本缺失或≤0，无法证明可盈利，不建议入盘。"
        elif current is None or current <= 0:
            reason = "商品池售价缺失或≤0，无法测算收益，不建议入盘。"
        elif profit_current <= 0:
            reason = "现池售价扣除3元履约成本后不盈利，不建议入盘。"
        elif profit_current < TARGET_PROFIT:
            reason = "现池售价可盈利，但低于3元目标单均收益，建议谨慎入盘。"
        elif profit_second is not None and profit_second <= 0:
            reason = "现池售价达到目标收益；但成本加2%会亏损，不能按加2%售价销售。"
        else:
            reason = "现池售价达到3元目标单均收益，且成本加2%仍可盈利。"
        stats[decision] += 1
        stats[f"5元:{decision5}"] += 1
        if profit_second is not None:
            stats["加2%盈利" if profit_second > 0 else "加2%不盈利"] += 1
        else:
            stats["加2%无法测算"] += 1

        def ref(name):
            return f"{get_column_letter(col[name])}{excel_row}"

        cost_ref = ref("采用成本(元)")
        current_ref = ref("商品池售价")
        fulfillment_ref = ref("一件代发履约成本(元/单)")
        second_ref = ref("一件代发加2%售价(元)")
        profit_current_ref = ref("现池售价单均收益(元)")
        formulas = [
            ("一件代发履约成本(元/单)", "='入盘参数'!$B$2", LOW_FULFILLMENT, money_fmt),
            ("一件代发加2%售价(元)", f'=IF(OR({cost_ref}="",{cost_ref}<=0),"",ROUND({cost_ref}*(1+\'入盘参数\'!$B$4),2))', second_price, money_fmt),
            ("加2%单均收益(元)", f'=IF({second_ref}="","",{second_ref}-{cost_ref}-{fulfillment_ref})', profit_second, money_fmt),
            ("现池售价单均收益(元)", f'=IF(OR({cost_ref}="",{cost_ref}<=0,{current_ref}="",{current_ref}<=0),"",{current_ref}-{cost_ref}-{fulfillment_ref})', profit_current, money_fmt),
            ("保本售价(元)", f'=IF(OR({cost_ref}="",{cost_ref}<=0),"",{cost_ref}+{fulfillment_ref})', break_even, money_fmt),
            ("目标售价(元)", f'=IF(OR({cost_ref}="",{cost_ref}<=0),"",{cost_ref}+{fulfillment_ref}+\'入盘参数\'!$B$3)', target_price, money_fmt),
            ("现售价与目标价差(元)", f'=IF(OR({current_ref}="",{current_ref}<=0,{ref("目标售价(元)")}=""),"",{current_ref}-{ref("目标售价(元)")})', target_diff, money_fmt),
            ("入盘建议", f'=IF(OR({cost_ref}="",{cost_ref}<=0,{current_ref}="",{current_ref}<=0),"不建议入盘",IF({profit_current_ref}>=\'入盘参数\'!$B$3,"建议入盘",IF({profit_current_ref}>0,"谨慎入盘","不建议入盘")))', decision, green_fmt if decision == "建议入盘" else yellow_fmt if decision == "谨慎入盘" else red_fmt),
            ("建议原因", "", reason, None),
            ("5元履约情景收益(元)", f'=IF(OR({cost_ref}="",{cost_ref}<=0,{current_ref}="",{current_ref}<=0),"",{current_ref}-{cost_ref}-\'入盘参数\'!$B$5)', profit5, money_fmt),
            ("5元情景建议", f'=IF(OR({cost_ref}="",{cost_ref}<=0,{current_ref}="",{current_ref}<=0),"不建议入盘",IF({ref("5元履约情景收益(元)")}>=\'入盘参数\'!$B$3,"建议入盘",IF({ref("5元履约情景收益(元)")}>0,"谨慎入盘","不建议入盘")))', decision5, green_fmt if decision5 == "建议入盘" else yellow_fmt if decision5 == "谨慎入盘" else red_fmt),
        ]
        # Reason is a deterministic explanatory label, not a numeric calculation; keep it as a formula-linked text result via IFs below.
        reason_formula = (
            f'=IF(OR({cost_ref}="",{cost_ref}<=0),"采用成本缺失或≤0，无法证明可盈利，不建议入盘。",'
            f'IF(OR({current_ref}="",{current_ref}<=0),"商品池售价缺失或≤0，无法测算收益，不建议入盘。",'
            f'IF({profit_current_ref}<=0,"现池售价扣除3元履约成本后不盈利，不建议入盘。",'
            f'IF({profit_current_ref}<\'入盘参数\'!$B$3,"现池售价可盈利，但低于3元目标单均收益，建议谨慎入盘。",'
            f'IF({ref("加2%单均收益(元)")}<=0,"现池售价达到目标收益；但成本加2%会亏损，不能按加2%售价销售。",'
            f'"现池售价达到3元目标单均收益，且成本加2%仍可盈利。")))))'
        )
        for name, formula, cached, fmt in formulas:
            if name == "建议原因":
                formula = reason_formula
            ws.write_formula(out_row, col[name] - 1, formula, fmt, "" if cached is None else cached)

    formula_wb.close()
    value_wb.close()

    advice_letter = get_column_letter(col["入盘建议"])
    second_profit_letter = get_column_letter(col["加2%单均收益(元)"])
    decision5_letter = get_column_letter(col["5元情景建议"])
    summary_metrics = [
        ("商品池商品总数", f"=COUNTA('二次定价'!$A$2:$A${max_row})", max_row - 1),
        ("建议入盘（3元履约）", f"=COUNTIF('二次定价'!${advice_letter}$2:${advice_letter}${max_row},\"建议入盘\")", stats["建议入盘"]),
        ("谨慎入盘（3元履约）", f"=COUNTIF('二次定价'!${advice_letter}$2:${advice_letter}${max_row},\"谨慎入盘\")", stats["谨慎入盘"]),
        ("不建议入盘（3元履约）", f"=COUNTIF('二次定价'!${advice_letter}$2:${advice_letter}${max_row},\"不建议入盘\")", stats["不建议入盘"]),
        ("成本加2%仍盈利", f"=COUNTIF('二次定价'!${second_profit_letter}$2:${second_profit_letter}${max_row},\">0\")", stats["加2%盈利"]),
        ("成本加2%不盈利", f"=COUNTIF('二次定价'!${second_profit_letter}$2:${second_profit_letter}${max_row},\"<=0\")", stats["加2%不盈利"]),
        ("建议入盘（5元履约）", f"=COUNTIF('二次定价'!${decision5_letter}$2:${decision5_letter}${max_row},\"建议入盘\")", stats["5元:建议入盘"]),
        ("谨慎入盘（5元履约）", f"=COUNTIF('二次定价'!${decision5_letter}$2:${decision5_letter}${max_row},\"谨慎入盘\")", stats["5元:谨慎入盘"]),
        ("不建议入盘（5元履约）", f"=COUNTIF('二次定价'!${decision5_letter}$2:${decision5_letter}${max_row},\"不建议入盘\")", stats["5元:不建议入盘"]),
    ]
    for r, (label, formula, cached) in enumerate(summary_metrics, start=7):
        summary.write(r, 0, label)
        summary.write_formula(r, 1, formula, int_fmt, cached)
    summary.write(18, 0, "执行建议", section_fmt)
    summary.write(18, 1, "说明", section_fmt)
    execution_notes = [
        ("建议入盘", "现池售价扣除3元履约成本后，单均收益≥3元；可进入货盘，但仍需过品质和供应商审核。"),
        ("谨慎入盘", "现池售价仅小幅盈利，收益不足3元；优先用于凑单、组合装或重新议价后再入。"),
        ("不建议入盘", "现池售价不能覆盖履约成本，或成本/售价缺失；不要仅靠低成本加价率判断。"),
        ("定价提醒", "成本加2%不盈利的商品，若按现池售价入盘，必须保留足够售价，不可切换成加2%低价。"),
        ("参数复核", "3元是低位假设；拿到实际快递、操作、包材和售后成本后，在“入盘参数”页替换并复算。"),
    ]
    for r, (label, text) in enumerate(execution_notes, start=19):
        summary.write(r, 0, label)
        summary.write(r, 1, text)

    workbook.close()
    repaired_escapes = repair_double_escaped_xml(OUTPUT)
    print(json.dumps({
        "output": str(OUTPUT), "products": max_row - 1, "columns": len(headers),
        "repaired_excel_escapes": repaired_escapes,
        "stats": stats,
        "parameters": {"low_fulfillment": LOW_FULFILLMENT, "normal_fulfillment": NORMAL_FULFILLMENT, "target_profit": TARGET_PROFIT, "markup": MARKUP},
    }, ensure_ascii=False, default=dict))


if __name__ == "__main__":
    main()
