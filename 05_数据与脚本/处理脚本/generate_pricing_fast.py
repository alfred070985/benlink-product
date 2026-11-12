#!/usr/bin/env python3
import json
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "商品池10.29.xlsx"
COSTS = ROOT / "online_product_costs.jsonl"
OUTPUT = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
CATEGORIES = [
    "日用百货", "粮油副食", "厨房用品", "软百", "休闲食品", "水饮冲调",
    "肉禽蛋品", "文体玩具", "熟食面点", "乳品母婴", "滋补保健", "水产海鲜",
    "水果", "3C数码", "酒类", "宠物食品及用品", "蔬菜", "鲜花",
]
PURCHASE_COST_CATEGORIES = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
COST_TYPE_PURCHASE = "采购成本（purchasePrice）"
COST_TYPE_WEIGHTED = "加权成本（costPrice）"


def safe(value):
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def number(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return value


def round2(value):
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def main():
    costs = {}
    with COSTS.open("r", encoding="utf-8") as f:
        for line in f:
            item = json.loads(line)
            costs[item["productId"]] = item

    source_wb = load_workbook(INPUT, read_only=True, data_only=True)
    source_ws = source_wb["DataMessage"]
    source_rows = source_ws.iter_rows(values_only=True)
    original_headers = list(next(source_rows))
    original_index = {name: i for i, name in enumerate(original_headers)}
    new_headers = [
        "商品系统编号", "主商品编号", "净重(g)", "发货方式编码", "线上发货方式",
        "加权成本costPrice(元)", "采购成本purchasePrice(元)", "来福池成本(元)",
        "成本取值口径", "采用成本(元)", "重量分档", "加价率", "二次定价(元)",
        "定价状态", "与原池售价差额(元)", "与原池售价差异率",
    ]
    headers = original_headers + new_headers
    col = {name: i + 1 for i, name in enumerate(headers)}
    max_row = source_ws.max_row

    workbook = xlsxwriter.Workbook(OUTPUT, {"constant_memory": True})
    workbook.set_calc_mode("auto")
    header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"})
    title_fmt = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E78"})
    section_fmt = workbook.add_format({"bold": True, "font_color": "#1F4E78"})
    input_fmt = workbook.add_format({"font_color": "#0000FF"})
    input_pct_fmt = workbook.add_format({"font_color": "#0000FF", "num_format": "0.00%"})
    external_fmt = workbook.add_format({"font_color": "#008000"})
    money_fmt = workbook.add_format({"num_format": "0.00"})
    pct_fmt = workbook.add_format({"num_format": "0.00%"})
    integer_fmt = workbook.add_format({"num_format": "#,##0"})
    note_fmt = workbook.add_format({"bg_color": "#FFF2CC", "text_wrap": True, "valign": "top"})

    summary = workbook.add_worksheet("汇总")
    summary.hide_gridlines(2)
    summary.set_column("A:A", 30)
    summary.set_column("B:B", 24)
    summary.set_column("D:D", 28)
    summary.set_column("E:E", 85)
    summary.write(0, 0, "来福商品池成本及二次定价", title_fmt)
    summary.write(1, 0, "源文件", section_fmt)
    summary.write(1, 1, "商品池10.29.xlsx")
    summary.write(2, 0, "成本数据环境", section_fmt)
    summary.write(2, 1, "online")
    summary.write(3, 0, "生成时间", section_fmt)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary.write(3, 1, generated_at)
    summary.write(5, 0, "结果指标", section_fmt)
    summary.write(5, 1, "商品数", section_fmt)

    params = workbook.add_worksheet("定价参数")
    params.hide_gridlines(2)
    params.set_column("A:A", 30)
    params.set_column("B:B", 18)
    params.set_column("D:D", 24)
    params.set_column("E:E", 32)
    params.set_column("G:G", 3)
    params.set_column("H:H", 100)
    params.write(0, 0, "定价参数（蓝色为可调整输入）", title_fmt)
    params.write(1, 0, "参数", section_fmt)
    params.write(1, 1, "值", section_fmt)
    params.write(2, 0, "云仓／代发加价率")
    params.write_number(2, 1, 0.02, input_pct_fmt)
    params.write(3, 0, "仓发≤重量阈值加价率")
    params.write_number(3, 1, 0.05, input_pct_fmt)
    params.write(4, 0, "仓发>重量阈值加价率")
    params.write_number(4, 1, 0.08, input_pct_fmt)
    params.write(5, 0, "重量阈值(g)")
    params.write_number(5, 1, 1000, input_fmt)
    params.write(7, 3, "商品大类", section_fmt)
    params.write(7, 4, "成本取值口径", section_fmt)
    category_start = 9
    for i, category in enumerate(CATEGORIES):
        cost_type = COST_TYPE_PURCHASE if category in PURCHASE_COST_CATEGORIES else COST_TYPE_WEIGHTED
        params.write(category_start - 1 + i, 3, category, input_fmt)
        params.write(category_start - 1 + i, 4, cost_type, input_fmt)
    params.write(27, 0, "执行口径", section_fmt)
    notes = [
        "原表“是否产地直发”的“代发”按用户所述“云仓”执行2%加价。",
        "1000g边界按“1000g（含）以下5%、1000g以上8%”执行。",
        "生鲜／水果初始映射为：水果、蔬菜、肉禽蛋品、水产海鲜、熟食面点；可在上方类目表调整。",
        "purchasePrice在商品主数据代码中的定义为默认采购价；本表将其作为采购成本口径字段，未声称已逐笔核验末次采购单。",
        "costPrice对应商品主数据单位成本，作为其他商品的加权成本口径。",
        "采用成本为空或≤0时不计算二次定价，定价状态标为“成本缺失”。",
    ]
    for i, note in enumerate(notes, start=28):
        params.write(i, 0, note, note_fmt)
    params.set_row(27, 30)

    sources = workbook.add_worksheet("数据来源")
    sources.hide_gridlines(2)
    sources.set_column("A:A", 38)
    sources.set_column("B:B", 105)
    sources.set_column("C:C", 85)
    sources.write_row(0, 0, ["Source Name", "Source URL", "用途与口径"], section_fmt)
    source_rows_data = [
        ["本来生活商品主数据 SOA（online）", "https://vm-online.benlai.com/business-product-servicehost/ProductApi/GetProductListByIds", "逐商品读取sysNo、productBasicSysNo、netWeight、deliveryMethod、costPrice、purchasePrice、lifeProductPoolCostPrice。"],
        ["商品池10.29.xlsx／DataMessage", "本地源文件", "保留原始31列；商品编号用于与线上成本数据一对一关联。"],
        ["线上代码：ProductService.cs", "benlai.business.product／master", "costPrice为单位成本；purchasePrice为默认采购价；lifeProductPoolCostPrice按发货方式生成。"],
        ["数据提取时间", generated_at, f"成本记录{len(costs):,}条；源商品{max_row - 1:,}条。"],
    ]
    for r, row in enumerate(source_rows_data, start=1):
        for c, value in enumerate(row):
            sources.write_string(r, c, str(value))

    ws = workbook.add_worksheet("二次定价")
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max_row - 1, len(headers) - 1)
    ws.set_row(0, 30)
    for c, name in enumerate(headers):
        ws.write(0, c, name, header_fmt)
    widths = {
        "商品编号": 16, "商品名称": 42, "商品大类": 16, "商品中类": 18, "商品小类": 20,
        "商品池售价": 14, "价格下限": 14, "价格上限": 14, "是否产地直发": 16,
        "净重(g)": 13, "成本取值口径": 28, "采用成本(元)": 16, "重量分档": 16,
        "加价率": 12, "二次定价(元)": 16, "定价状态": 20, "与原池售价差异率": 20,
    }
    for name, idx in col.items():
        ws.set_column(idx - 1, idx - 1, widths.get(name, 15))

    numeric_original = {"长期售价", "销项税率", "商品池售价", "价格下限", "价格上限", "分销渠道在线库存", "近30天销量(出库数量)", "C网零售价", "京东商品价格"}
    stats = {"已定价": 0, "成本缺失": 0, "发货方式或重量缺失": 0, "云仓": 0, "仓发": 0, COST_TYPE_PURCHASE: 0, COST_TYPE_WEIGHTED: 0}
    for excel_row, row in enumerate(source_rows, start=2):
        out_row = excel_row - 1
        values = list(row)
        product_id = str(values[original_index["商品编号"]]).strip()
        item = costs.get(product_id)
        if item is None:
            raise RuntimeError(f"No online cost record for {product_id}")
        for name in numeric_original:
            values[original_index[name]] = number(values[original_index[name]])
        for c, value in enumerate(values):
            value = safe(value)
            if value is not None:
                ws.write(out_row, c, value)

        delivery_code = item.get("deliveryMethod")
        delivery_name = {1: "仓发", 2: "代发／云仓"}.get(delivery_code, "未知")
        external_values = [
            item.get("sysNo"), item.get("productBasicSysNo"), item.get("netWeight"), delivery_code, delivery_name,
            number(item.get("costPrice")), number(item.get("purchasePrice")), number(item.get("lifeProductPoolCostPrice")),
        ]
        for offset, value in enumerate(external_values):
            if value is not None:
                ws.write(out_row, 31 + offset, value, external_fmt)

        category = values[original_index["商品大类"]]
        original_delivery = values[original_index["是否产地直发"]]
        weight = item.get("netWeight")
        weighted = number(item.get("costPrice"))
        purchase = number(item.get("purchasePrice"))
        cost_type = COST_TYPE_PURCHASE if category in PURCHASE_COST_CATEGORIES else COST_TYPE_WEIGHTED
        selected = purchase if cost_type == COST_TYPE_PURCHASE else weighted
        if original_delivery in ("云仓", "代发"):
            rate = 0.02
            band = "不适用"
            stats["云仓"] += 1
        elif original_delivery == "仓发":
            stats["仓发"] += 1
            if weight is None:
                rate = None
                band = "重量缺失"
            else:
                rate = 0.05 if weight <= 1000 else 0.08
                band = "≤阈值" if weight <= 1000 else ">阈值"
        else:
            rate = None
            band = "不适用"
        stats[cost_type] += 1
        new_price = round2(selected * (1 + rate)) if selected is not None and selected > 0 and rate is not None else None
        if selected is None or selected <= 0:
            status = "成本缺失"
        elif rate is None:
            status = "发货方式或重量缺失"
        else:
            status = "已定价"
        stats[status] += 1
        pool_price = values[original_index["商品池售价"]]
        diff = new_price - pool_price if new_price is not None and isinstance(pool_price, (int, float)) else None
        diff_rate = diff / pool_price if diff is not None and pool_price else None

        def ref(name):
            return f"{get_column_letter(col[name])}{excel_row}"

        category_ref = ref("商品大类")
        original_delivery_ref = ref("是否产地直发")
        weight_ref = ref("净重(g)")
        weighted_ref = ref("加权成本costPrice(元)")
        purchase_ref = ref("采购成本purchasePrice(元)")
        cost_type_ref = ref("成本取值口径")
        selected_cost_ref = ref("采用成本(元)")
        rate_ref = ref("加价率")
        new_price_ref = ref("二次定价(元)")
        pool_price_ref = ref("商品池售价")
        diff_ref = ref("与原池售价差额(元)")
        category_lookup = f"IFERROR(INDEX('定价参数'!$E${category_start}:$E${category_start + len(CATEGORIES) - 1},MATCH({category_ref},'定价参数'!$D${category_start}:$D${category_start + len(CATEGORIES) - 1},0)),\"{COST_TYPE_WEIGHTED}\")"
        formulas = [
            ("成本取值口径", f"={category_lookup}", cost_type, None),
            ("采用成本(元)", f'=IF({cost_type_ref}="{COST_TYPE_PURCHASE}",{purchase_ref},{weighted_ref})', selected, money_fmt),
            ("重量分档", f'=IF(OR({original_delivery_ref}="云仓",{original_delivery_ref}="代发"),"不适用",IF({weight_ref}="","重量缺失",IF({weight_ref}<=\'定价参数\'!$B$6,"≤阈值",">阈值")))', band, None),
            ("加价率", f'=IF(OR({original_delivery_ref}="云仓",{original_delivery_ref}="代发"),\'定价参数\'!$B$3,IF({original_delivery_ref}="仓发",IF({weight_ref}="","",IF({weight_ref}<=\'定价参数\'!$B$6,\'定价参数\'!$B$4,\'定价参数\'!$B$5)),""))', rate, pct_fmt),
            ("二次定价(元)", f'=IF(OR({selected_cost_ref}="",{selected_cost_ref}<=0,{rate_ref}=""),"",ROUND({selected_cost_ref}*(1+{rate_ref}),2))', new_price, money_fmt),
            ("定价状态", f'=IF(OR({selected_cost_ref}="",{selected_cost_ref}<=0),"成本缺失",IF({rate_ref}="","发货方式或重量缺失","已定价"))', status, None),
            ("与原池售价差额(元)", f'=IF(OR({new_price_ref}="",{pool_price_ref}=""),"",{new_price_ref}-{pool_price_ref})', diff, money_fmt),
            ("与原池售价差异率", f'=IF(OR({new_price_ref}="",{pool_price_ref}="",{pool_price_ref}=0),"",{diff_ref}/{pool_price_ref})', diff_rate, pct_fmt),
        ]
        for name, formula, cached, fmt in formulas:
            ws.write_formula(out_row, col[name] - 1, formula, fmt, "" if cached is None else cached)

    source_wb.close()
    status_letter = get_column_letter(col["定价状态"])
    delivery_letter = get_column_letter(col["是否产地直发"])
    cost_type_letter = get_column_letter(col["成本取值口径"])
    summary_metrics = [
        ("商品池商品总数", f"=COUNTA('二次定价'!$A$2:$A${max_row})", max_row - 1),
        ("已定价", f"=COUNTIF('二次定价'!${status_letter}$2:${status_letter}${max_row},\"已定价\")", stats["已定价"]),
        ("成本缺失", f"=COUNTIF('二次定价'!${status_letter}$2:${status_letter}${max_row},\"成本缺失\")", stats["成本缺失"]),
        ("云仓／代发商品", f"=COUNTIF('二次定价'!${delivery_letter}$2:${delivery_letter}${max_row},\"代发\")+COUNTIF('二次定价'!${delivery_letter}$2:${delivery_letter}${max_row},\"云仓\")", stats["云仓"]),
        ("仓发商品", f"=COUNTIF('二次定价'!${delivery_letter}$2:${delivery_letter}${max_row},\"仓发\")", stats["仓发"]),
        ("采购成本口径商品", f"=COUNTIF('二次定价'!${cost_type_letter}$2:${cost_type_letter}${max_row},\"{COST_TYPE_PURCHASE}\")", stats[COST_TYPE_PURCHASE]),
        ("加权成本口径商品", f"=COUNTIF('二次定价'!${cost_type_letter}$2:${cost_type_letter}${max_row},\"{COST_TYPE_WEIGHTED}\")", stats[COST_TYPE_WEIGHTED]),
    ]
    for r, (label, formula, cached) in enumerate(summary_metrics, start=6):
        summary.write(r, 0, label)
        summary.write_formula(r, 1, formula, integer_fmt, cached)
    summary.write(14, 0, "定价规则", section_fmt)
    summary.write(14, 1, "口径说明", section_fmt)
    rules = [
        ("云仓", "原表“代发”按云仓处理：采用成本 ×（1＋2%）。"),
        ("仓发", "净重≤1000g：采用成本 ×（1＋5%）；净重>1000g：采用成本 ×（1＋8%）。"),
        ("成本选择", "水果、蔬菜、肉禽蛋品、水产海鲜、熟食面点使用purchasePrice；其他类目使用costPrice。"),
        ("采购成本字段提示", "线上商品主数据的purchasePrice代码定义为默认采购价；本工作簿按指定采购成本口径使用该字段，未逐笔追溯采购单。"),
        ("缺失成本", "采用成本为空或≤0时不计算二次定价，定价状态标为“成本缺失”，两种原始成本仍保留供核查。"),
        ("价格精度", "二次定价四舍五入保留2位小数；所有计算列均为Excel公式，参数可在“定价参数”页调整。"),
    ]
    for r, (label, text) in enumerate(rules, start=15):
        summary.write(r, 0, label)
        summary.write(r, 1, text)
    summary.write(22, 0, "工作表", section_fmt)
    summary.write(22, 1, "内容", section_fmt)
    sheet_index = [
        ("二次定价", "原商品池31列＋成本、重量、口径、加价率、二次定价及差异结果。"),
        ("定价参数", "加价率、重量阈值、类目成本口径与执行假设。"),
        ("数据来源", "线上接口、源文件和字段映射说明。"),
    ]
    for r, (label, text) in enumerate(sheet_index, start=23):
        summary.write(r, 0, label)
        summary.write(r, 1, text)

    workbook.close()
    print(json.dumps({"output": str(OUTPUT), "products": max_row - 1, "cost_records": len(costs), "columns": len(headers), "stats": stats}, ensure_ascii=False))


if __name__ == "__main__":
    main()
