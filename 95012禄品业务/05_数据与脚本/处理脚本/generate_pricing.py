#!/usr/bin/env python3
import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
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

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
TITLE_FONT = Font(color="1F4E78", bold=True, size=16)
SECTION_FONT = Font(color="1F4E78", bold=True, size=11)
INPUT_FONT = Font(color="0000FF")
FORMULA_FONT = Font(color="000000")
EXTERNAL_FONT = Font(color="008000")
LIGHT_FILL = PatternFill("solid", fgColor="D9EAF7")
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")


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


def styled_cell(ws, value, *, font=None, fill=None, num_format=None, align=None):
    cell = WriteOnlyCell(ws, value=safe(value))
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if num_format:
        cell.number_format = num_format
    if align:
        cell.alignment = align
    return cell


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

    out = Workbook(write_only=True)
    try:
        out.calculation.fullCalcOnLoad = True
        out.calculation.forceFullCalc = True
        out.calculation.calcMode = "auto"
    except AttributeError:
        pass

    summary = out.create_sheet("汇总")
    summary.sheet_view.showGridLines = False
    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 22
    summary.column_dimensions["D"].width = 28
    summary.column_dimensions["E"].width = 80
    summary.append([styled_cell(summary, "来福商品池成本及二次定价", font=TITLE_FONT)])
    summary.append([styled_cell(summary, "源文件", font=SECTION_FONT), "商品池10.29.xlsx"])
    summary.append([styled_cell(summary, "成本数据环境", font=SECTION_FONT), "online"])
    summary.append([styled_cell(summary, "生成时间", font=SECTION_FONT), datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    summary.append([])
    summary.append([styled_cell(summary, "结果指标", font=SECTION_FONT), styled_cell(summary, "商品数", font=SECTION_FONT)])
    ws = out.create_sheet("二次定价")
    max_row = source_ws.max_row
    status_col = get_column_letter(col["定价状态"])
    delivery_original_col = get_column_letter(col["是否产地直发"])
    cost_type_col = get_column_letter(col["成本取值口径"])
    cloud_count = f'COUNTIF(\'二次定价\'!${delivery_original_col}$2:${delivery_original_col}${max_row},"代发")+COUNTIF(\'二次定价\'!${delivery_original_col}$2:${delivery_original_col}${max_row},"云仓")'
    summary_metrics = [
        ("商品池商品总数", f"=COUNTA('二次定价'!$A$2:$A${max_row})"),
        ("已定价", f"=COUNTIF('二次定价'!${status_col}$2:${status_col}${max_row},\"已定价\")"),
        ("成本缺失", f"=COUNTIF('二次定价'!${status_col}$2:${status_col}${max_row},\"成本缺失\")"),
        ("云仓／代发商品", f"={cloud_count}"),
        ("仓发商品", f"=COUNTIF('二次定价'!${delivery_original_col}$2:${delivery_original_col}${max_row},\"仓发\")"),
        ("采购成本口径商品", f"=COUNTIF('二次定价'!${cost_type_col}$2:${cost_type_col}${max_row},\"{COST_TYPE_PURCHASE}\")"),
        ("加权成本口径商品", f"=COUNTIF('二次定价'!${cost_type_col}$2:${cost_type_col}${max_row},\"{COST_TYPE_WEIGHTED}\")"),
    ]
    for label, formula in summary_metrics:
        summary.append([label, styled_cell(summary, formula, font=FORMULA_FONT, num_format="#,##0")])
    summary.append([])
    summary.append([styled_cell(summary, "定价规则", font=SECTION_FONT), styled_cell(summary, "口径说明", font=SECTION_FONT)])
    rules = [
        ("云仓", "原表“代发”按云仓处理：采用成本 ×（1＋2%）。"),
        ("仓发", "净重≤1000g：采用成本 ×（1＋5%）；净重>1000g：采用成本 ×（1＋8%）。"),
        ("成本选择", "水果、蔬菜、肉禽蛋品、水产海鲜、熟食面点使用 purchasePrice；其他类目使用 costPrice。"),
        ("采购成本字段提示", "线上商品主数据的 purchasePrice 代码定义为默认采购价；本工作簿按指定采购成本口径使用该字段，未逐笔追溯采购单。"),
        ("缺失成本", "采用成本为空或≤0时不计算二次定价，定价状态标为“成本缺失”，两种原始成本仍保留供核查。"),
        ("价格精度", "二次定价四舍五入保留2位小数；所有计算列均为Excel公式，参数可在“定价参数”页调整。"),
    ]
    for label, text in rules:
        summary.append([label, text])
    summary.append([])
    summary.append([styled_cell(summary, "工作表", font=SECTION_FONT), styled_cell(summary, "内容", font=SECTION_FONT)])
    summary.append(["二次定价", "原商品池31列＋成本、重量、口径、加价率、二次定价及差异结果。"])
    summary.append(["定价参数", "加价率、重量阈值、类目成本口径与执行假设。"])
    summary.append(["数据来源", "线上接口、源文件和字段映射说明。"])

    params = out.create_sheet("定价参数")
    params.sheet_view.showGridLines = False
    for letter, width in {"A": 30, "B": 18, "D": 24, "E": 32, "G": 24, "H": 95}.items():
        params.column_dimensions[letter].width = width
    params.append([styled_cell(params, "定价参数（蓝色为可调整输入）", font=TITLE_FONT)])
    params.append([styled_cell(params, "参数", font=SECTION_FONT), styled_cell(params, "值", font=SECTION_FONT)])
    params.append(["云仓／代发加价率", styled_cell(params, 0.02, font=INPUT_FONT, num_format="0.00%")])
    params.append(["仓发≤重量阈值加价率", styled_cell(params, 0.05, font=INPUT_FONT, num_format="0.00%")])
    params.append(["仓发>重量阈值加价率", styled_cell(params, 0.08, font=INPUT_FONT, num_format="0.00%")])
    params.append(["重量阈值(g)", styled_cell(params, 1000, font=INPUT_FONT, num_format="0")])
    params.append([])
    params.append([styled_cell(params, "商品大类", font=SECTION_FONT), styled_cell(params, "成本取值口径", font=SECTION_FONT), None,
                   styled_cell(params, "商品大类", font=SECTION_FONT), styled_cell(params, "成本取值口径", font=SECTION_FONT)])
    category_start = 9
    for i, category in enumerate(CATEGORIES):
        row = category_start + i
        cost_type = COST_TYPE_PURCHASE if category in PURCHASE_COST_CATEGORIES else COST_TYPE_WEIGHTED
        # Categories occupy D:E so the main-sheet formulas can use INDEX/MATCH.
        while params._current_row < row - 1:
            params.append([])
        params.append([None, None, None,
                       styled_cell(params, category, font=INPUT_FONT),
                       styled_cell(params, cost_type, font=INPUT_FONT)])
    params.append([])
    params.append([styled_cell(params, "执行口径", font=SECTION_FONT), styled_cell(params, "说明", font=SECTION_FONT)])
    notes = [
        "原表“是否产地直发”的“代发”按用户所述“云仓”执行2%加价。",
        "1000g边界按“1000g（含）以下5%、1000g以上8%”执行。",
        "生鲜／水果初始映射为：水果、蔬菜、肉禽蛋品、水产海鲜、熟食面点；可在上方类目表调整。",
        "purchasePrice在商品主数据代码中的定义为默认采购价；本表将其作为采购成本口径字段，未声称已逐笔核验末次采购单。",
        "costPrice对应商品主数据单位成本，作为其他商品的加权成本口径。",
    ]
    for note in notes:
        params.append([None, None, None, None, None, None, styled_cell(params, note, fill=NOTE_FILL)])

    sources = out.create_sheet("数据来源")
    sources.sheet_view.showGridLines = False
    sources.column_dimensions["A"].width = 38
    sources.column_dimensions["B"].width = 105
    sources.column_dimensions["C"].width = 85
    sources.append([styled_cell(sources, "Source Name", font=SECTION_FONT), styled_cell(sources, "Source URL", font=SECTION_FONT), styled_cell(sources, "用途与口径", font=SECTION_FONT)])
    sources.append([
        "本来生活商品主数据 SOA（online）",
        "https://vm-online.benlai.com/business-product-servicehost/ProductApi/GetProductListByIds",
        "逐商品读取sysNo、productBasicSysNo、netWeight、deliveryMethod、costPrice、purchasePrice、lifeProductPoolCostPrice。",
    ])
    sources.append(["商品池10.29.xlsx／DataMessage", "本地源文件", "保留原始31列；商品编号用于与线上成本数据一对一关联。"])
    sources.append(["线上代码：ProductService.cs", "benlai.business.product／master", "costPrice为单位成本；purchasePrice为默认采购价；lifeProductPoolCostPrice按发货方式生成。"])
    sources.append(["数据提取时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), f"成本记录{len(costs):,}条；源商品{max_row - 1:,}条。"])

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max_row}"
    header_cells = [styled_cell(ws, h, font=HEADER_FONT, fill=HEADER_FILL, align=Alignment(horizontal="center", vertical="center")) for h in headers]
    ws.append(header_cells)
    widths = {
        "商品编号": 16, "商品名称": 42, "商品大类": 16, "商品中类": 18, "商品小类": 20,
        "商品池售价": 14, "价格下限": 14, "价格上限": 14, "是否产地直发": 16,
        "净重(g)": 13, "成本取值口径": 28, "采用成本(元)": 16, "重量分档": 16,
        "加价率": 12, "二次定价(元)": 16, "定价状态": 18,
    }
    for name, idx in col.items():
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(name, 15)

    numeric_original = {"长期售价", "销项税率", "商品池售价", "价格下限", "价格上限", "分销渠道在线库存", "近30天销量(出库数量)", "C网零售价", "京东商品价格"}
    missing_products = 0
    for excel_row, row in enumerate(source_rows, start=2):
        values = list(row)
        product_id = str(values[original_index["商品编号"]]).strip()
        item = costs.get(product_id)
        if item is None:
            missing_products += 1
            item = {}
        for name in numeric_original:
            values[original_index[name]] = number(values[original_index[name]])

        delivery_code = item.get("deliveryMethod")
        delivery_name = {1: "仓发", 2: "代发／云仓"}.get(delivery_code, "未知")
        external_values = [
            item.get("sysNo"), item.get("productBasicSysNo"), item.get("netWeight"), delivery_code, delivery_name,
            number(item.get("costPrice")), number(item.get("purchasePrice")), number(item.get("lifeProductPoolCostPrice")),
        ]
        cells = [safe(v) for v in values]
        cells.extend(external_values)

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
        cost_type_formula = f"={category_lookup}"
        selected_cost_formula = f'=IF({cost_type_ref}="{COST_TYPE_PURCHASE}",{purchase_ref},{weighted_ref})'
        weight_band_formula = f'=IF(OR({original_delivery_ref}="云仓",{original_delivery_ref}="代发"),"不适用",IF({weight_ref}="","重量缺失",IF({weight_ref}<=\'定价参数\'!$B$6,"≤阈值",">阈值")))'
        rate_formula = f'=IF(OR({original_delivery_ref}="云仓",{original_delivery_ref}="代发"),\'定价参数\'!$B$3,IF({original_delivery_ref}="仓发",IF({weight_ref}="","",IF({weight_ref}<=\'定价参数\'!$B$6,\'定价参数\'!$B$4,\'定价参数\'!$B$5)),""))'
        new_price_formula = f'=IF(OR({selected_cost_ref}="",{selected_cost_ref}<=0,{rate_ref}=""),"",ROUND({selected_cost_ref}*(1+{rate_ref}),2))'
        status_formula = f'=IF(OR({selected_cost_ref}="",{selected_cost_ref}<=0),"成本缺失",IF({rate_ref}="","发货方式或重量缺失","已定价"))'
        diff_formula = f'=IF(OR({new_price_ref}="",{pool_price_ref}=""),"",{new_price_ref}-{pool_price_ref})'
        diff_rate_formula = f'=IF(OR({new_price_ref}="",{pool_price_ref}="",{pool_price_ref}=0),"",{diff_ref}/{pool_price_ref})'
        formula_values = [cost_type_formula, selected_cost_formula, weight_band_formula, rate_formula, new_price_formula, status_formula, diff_formula, diff_rate_formula]
        for name, formula in zip(new_headers[8:], formula_values):
            fmt = None
            if name in {"采用成本(元)", "二次定价(元)", "与原池售价差额(元)"}:
                fmt = "0.00"
            elif name in {"加价率", "与原池售价差异率"}:
                fmt = "0.00%"
            cells.append(styled_cell(ws, formula, font=FORMULA_FONT, num_format=fmt))
        ws.append(cells)

    source_wb.close()
    if missing_products:
        raise RuntimeError(f"{missing_products} products have no online cost record")
    out.save(OUTPUT)
    print(json.dumps({"output": str(OUTPUT), "products": max_row - 1, "cost_records": len(costs), "columns": len(headers)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
