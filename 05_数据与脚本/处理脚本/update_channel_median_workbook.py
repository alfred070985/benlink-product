#!/usr/bin/env python3
import json
from collections import Counter
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "来福商品池_成本及二次定价_末次采购成本更新_2026-10-30.xlsx"
ERP = ROOT / "product_channel_median_costs.jsonl"
PROGRESS = ROOT / "channel_price_fetch_progress.json"
LATEST = ROOT / "latest_purchase_costs_2026-10-01_2026-10-30.jsonl"
PURCHASE_PROGRESS = ROOT / "latest_purchase_costs_progress.json"
COUNT_JSON = ROOT / "purchase_callbacks_count.json"
OUTPUT = ROOT / "来福商品池_成本及二次定价_渠道中位数加权成本_2026-10-30.xlsx"
SPECIAL = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
PURCHASE = "末次采购成本（采购明细）"
OLD_WEIGHTED = "加权成本（costPrice）"
WEIGHTED = "加权成本（分销/C网/非集采中位数）"
OLD_WEIGHTED_HEADER = "加权成本costPrice(元)"
WEIGHTED_HEADER = "加权成本UnitCost中位数(元)"
CALCULATED = {"成本取值口径", "采用成本(元)", "重量分档", "加价率", "二次定价(元)", "定价状态"}


def number(value):
    if value is None or value == "":
        return None
    return float(value)


def round2(value):
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def safe(value):
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def load_jsonl(path, key):
    rows = {}
    with path.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                identity = item[key]
                if identity in rows:
                    raise RuntimeError(f"Duplicate {key}: {identity}")
                rows[identity] = item
    return rows


def main():
    progress = json.loads(PROGRESS.read_text(encoding="utf-8"))
    purchase_progress = json.loads(PURCHASE_PROGRESS.read_text(encoding="utf-8"))
    if not progress.get("coverage_complete") or not purchase_progress.get("coverage_complete"):
        raise RuntimeError("Cost collection is incomplete")
    erp = load_jsonl(ERP, "product_id")
    latest = load_jsonl(LATEST, "product_id")
    if any(item.get("UnitCost") is None for item in erp.values()):
        raise RuntimeError("Null UnitCost requires a separate missing-data rule")
    raw_detail_count = json.loads(COUNT_JSON.read_text(encoding="utf-8"))["data"]["rows"][0][0]
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    source_formula = load_workbook(INPUT, read_only=True, data_only=False)
    source_values = load_workbook(INPUT, read_only=True, data_only=True)
    small_formula = {}
    for name in ("定价参数", "数据来源"):
        small_formula[name] = list(source_formula[name].iter_rows(values_only=True))
    sf = source_formula["二次定价"]
    sv = source_values["二次定价"]
    formula_iter = sf.iter_rows(values_only=True)
    value_iter = sv.iter_rows(values_only=True)
    source_headers = list(next(formula_iter))
    cached_headers = list(next(value_iter))
    if source_headers != cached_headers or len(source_headers) != 31:
        raise RuntimeError("Unexpected source workbook structure")
    col = {name: idx for idx, name in enumerate(source_headers)}
    required = ["商品编号", "商品大类", "是否产地直发", "净重(g)", OLD_WEIGHTED_HEADER,
                "末次采购成本(元)", "成本取值口径", "采用成本(元)", "重量分档", "加价率", "二次定价(元)", "定价状态"]
    if any(name not in col for name in required):
        raise RuntimeError("Required source columns are missing")
    headers = [WEIGHTED_HEADER if name == OLD_WEIGHTED_HEADER else name for name in source_headers]

    workbook = xlsxwriter.Workbook(OUTPUT, {"constant_memory": True})
    workbook.set_calc_mode("auto")
    header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"})
    title_fmt = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E78"})
    section_fmt = workbook.add_format({"bold": True, "font_color": "#1F4E78"})
    input_fmt = workbook.add_format({"font_color": "#0000FF"})
    input_pct_fmt = workbook.add_format({"font_color": "#0000FF", "num_format": "0.00%"})
    external_fmt = workbook.add_format({"font_color": "#008000"})
    external_cost_fmt = workbook.add_format({"font_color": "#008000", "num_format": "0.000000"})
    cost_fmt = workbook.add_format({"num_format": "0.000000"})
    money_fmt = workbook.add_format({"num_format": "0.00"})
    pct_fmt = workbook.add_format({"num_format": "0.00%"})
    integer_fmt = workbook.add_format({"num_format": "#,##0"})
    note_fmt = workbook.add_format({"bg_color": "#FFF2CC", "text_wrap": True, "valign": "top"})
    qty_fmt = workbook.add_format({"num_format": "0.0000"})
    tax_fmt = workbook.add_format({"num_format": "0.000000"})

    summary = workbook.add_worksheet("汇总")
    params = workbook.add_worksheet("定价参数")
    sources = workbook.add_worksheet("数据来源")
    ws = workbook.add_worksheet("二次定价")
    detail = workbook.add_worksheet("末次采购明细")
    weighted_detail = workbook.add_worksheet("加权成本取价明细")

    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, sf.max_row - 1, len(headers) - 1)
    ws.set_row(0, 30)
    for c, name in enumerate(headers):
        ws.write(0, c, name, header_fmt)
    widths = {
        "商品编号": 16, "商品名称": 42, "商品大类": 16, "商品中类": 18, "商品小类": 20,
        "商品池售价": 14, "价格下限": 14, "价格上限": 14, "是否产地直发": 16,
        "净重(g)": 13, "成本取值口径": 36, "采用成本(元)": 16, "重量分档": 16,
        "加价率": 12, "二次定价(元)": 16, "定价状态": 24, "末次采购成本(元)": 22,
        WEIGHTED_HEADER: 24,
    }
    for name, idx in col.items():
        output_name = WEIGHTED_HEADER if name == OLD_WEIGHTED_HEADER else name
        ws.set_column(idx, idx, widths.get(output_name, 15))

    stats = Counter()
    row_count = 0
    pool_ids = set()
    for excel_row, (formula_row, cached_row) in enumerate(zip(formula_iter, value_iter), start=2):
        row_count += 1
        out_row = excel_row - 1
        values = list(formula_row)
        cached = list(cached_row)
        pid = str(values[col["商品编号"]]).strip()
        pool_ids.add(pid)
        erp_item = erp.get(pid)
        weighted = number(erp_item["UnitCost"]) if erp_item is not None else None
        purchase = number(cached[col["末次采购成本(元)"]])
        stats["有Product_Price记录" if erp_item is not None else "无Product_Price记录"] += 1

        category = values[col["商品大类"]]
        delivery = values[col["是否产地直发"]]
        weight = number(cached[col["净重(g)"]])
        cost_type = PURCHASE if category in SPECIAL else WEIGHTED
        selected = purchase if cost_type == PURCHASE else weighted
        if delivery in ("云仓", "代发"):
            rate = 0.02
            band = "不适用"
            stats["云仓／代发商品"] += 1
        elif delivery == "仓发":
            stats["仓发商品"] += 1
            if weight is None:
                rate = None
                band = "重量缺失"
            else:
                rate = 0.05 if weight <= 1000 else 0.08
                band = "≤阈值" if weight <= 1000 else ">阈值"
        else:
            rate = None
            band = "不适用"
        new_price = round2(selected * (1 + rate)) if selected is not None and selected > 0 and rate is not None else None
        if cost_type == PURCHASE and purchase is None:
            status = "本期无采购记录"
        elif selected is None or selected <= 0:
            status = "成本缺失"
        elif rate is None:
            status = "发货方式或重量缺失"
        else:
            status = "已定价"
        stats[cost_type] += 1
        stats[status] += 1

        formula_cache = {
            "成本取值口径": cost_type,
            "采用成本(元)": selected,
            "重量分档": band,
            "加价率": rate,
            "二次定价(元)": new_price,
            "定价状态": status,
        }
        formula_formats = {
            "成本取值口径": None,
            "采用成本(元)": cost_fmt,
            "重量分档": None,
            "加价率": pct_fmt,
            "二次定价(元)": money_fmt,
            "定价状态": None,
        }
        for idx, source_name in enumerate(source_headers):
            value = values[idx]
            if source_name in CALCULATED:
                formula = value.replace(OLD_WEIGHTED, WEIGHTED) if isinstance(value, str) else value
                if not (isinstance(formula, str) and formula.startswith("=")):
                    raise RuntimeError(f"Missing formula at {source_name}, row {excel_row}")
                cached_value = formula_cache[source_name]
                ws.write_formula(out_row, idx, formula, formula_formats[source_name], "" if cached_value is None else cached_value)
            elif source_name == OLD_WEIGHTED_HEADER:
                if weighted is not None:
                    ws.write_number(out_row, idx, weighted, external_cost_fmt)
            elif value is not None:
                if source_name == "末次采购成本(元)":
                    ws.write_number(out_row, idx, number(value), external_cost_fmt)
                elif 17 <= idx <= 24:
                    ws.write(out_row, idx, safe(value), external_fmt)
                else:
                    ws.write(out_row, idx, safe(value))
    if sum(1 for _ in formula_iter) or sum(1 for _ in value_iter) or row_count != 43519:
        raise RuntimeError("Source row count mismatch")
    if len(pool_ids) != row_count or not set(erp).issubset(pool_ids) or not set(latest).issubset(pool_ids):
        raise RuntimeError("Product identity mismatch")

    params.hide_gridlines(2)
    params.set_column("A:A", 110)
    params.set_column("B:B", 18)
    params.set_column("D:D", 24)
    params.set_column("E:E", 40)
    for r, row in enumerate(small_formula["定价参数"]):
        for c, original in enumerate(row):
            value = original
            if isinstance(value, str):
                value = value.replace(OLD_WEIGHTED, WEIGHTED)
                if "costPrice对应商品主数据单位成本" in value:
                    value = "非生鲜／水果商品按ProductBasicSysNo关联多渠道Product_Price.UnitCost：分销→C网→排除集采后的全部渠道，命中层级多条取中位数；不再使用商品主数据costPrice。"
                elif "采用成本为空或≤0时" in value:
                    value = "采用成本为空或≤0时不计算二次定价；无有效渠道价格记录时加权成本留空并标记“成本缺失”；生鲜／水果无本期采购记录时标记“本期无采购记录”。"
            if value is None:
                continue
            fmt = None
            if r == 0:
                fmt = title_fmt
            elif r in (1, 7, 27):
                fmt = section_fmt
            elif r in (2, 3, 4, 5) and c == 1:
                fmt = input_pct_fmt if r in (2, 3, 4) else input_fmt
            elif 8 <= r <= 25 and c in (3, 4):
                fmt = input_fmt
            elif r >= 28 and c == 0:
                fmt = note_fmt
            params.write(r, c, value, fmt)
    params.set_row(28, 32)
    params.set_row(31, 54)
    params.set_row(32, 54)
    params.set_row(33, 48)

    detail_headers = [
        "采购时间(inventory_time)", "回调系统编号(callback_sys_no)", "回调单号(callback_order_code)",
        "仓库编码(warehouse_code)", "单据类型(doc_type)", "商品MDM编码(item_code)",
        "商品编号(product_id)", "商品系统编号(product_sysno)", "数量(qty)",
        "末次采购成本(buy_price)", "税率(tax_rate)",
    ]
    detail_keys = ["inventory_time", "callback_sys_no", "callback_order_code", "warehouse_code", "doc_type", "item_code", "product_id", "product_sysno", "qty", "buy_price", "tax_rate"]
    detail.hide_gridlines(2)
    detail.freeze_panes(1, 0)
    detail.autofilter(0, 0, len(latest), len(detail_headers) - 1)
    detail.set_row(0, 36)
    detail.set_column("A:A", 22)
    detail.set_column("B:B", 24)
    detail.set_column("C:C", 28)
    detail.set_column("D:E", 18)
    detail.set_column("F:F", 22)
    detail.set_column("G:H", 26)
    detail.set_column("I:K", 24)
    for c, name in enumerate(detail_headers):
        detail.write(0, c, name, header_fmt)
    for r, pid in enumerate(sorted(latest), start=1):
        item = latest[pid]
        for c, key in enumerate(detail_keys):
            value = item.get(key)
            if value is None:
                continue
            if key in ("qty", "buy_price", "tax_rate"):
                fmt = qty_fmt if key == "qty" else cost_fmt if key == "buy_price" else tax_fmt
                detail.write_number(r, c, float(value), fmt)
            elif key in ("callback_sys_no", "item_code"):
                detail.write_number(r, c, float(value), integer_fmt)
            else:
                detail.write(r, c, safe(value))

    weighted_headers = [
        "商品编号", "主商品编号(ProductBasicSysNo)", "取价渠道", "中位数UnitCost(元)",
        "取价记录数", "站点数", "渠道商品SysNo列表", "站点列表",
        "最低UnitCost(元)", "最高UnitCost(元)", "中位数规则",
    ]
    weighted_detail.hide_gridlines(2)
    weighted_detail.freeze_panes(1, 0)
    weighted_detail.autofilter(0, 0, len(erp), len(weighted_headers) - 1)
    weighted_detail.set_row(0, 36)
    weighted_detail.set_column("A:A", 16)
    weighted_detail.set_column("B:B", 30)
    weighted_detail.set_column("C:C", 20)
    weighted_detail.set_column("D:D", 24)
    weighted_detail.set_column("E:F", 14)
    weighted_detail.set_column("G:G", 70)
    weighted_detail.set_column("H:H", 24)
    weighted_detail.set_column("I:J", 22)
    weighted_detail.set_column("K:K", 58)
    for c, name in enumerate(weighted_headers):
        weighted_detail.write(0, c, name, header_fmt)
    for r, pid in enumerate(sorted(erp), start=1):
        item = erp[pid]
        weighted_detail.write(r, 0, pid, external_fmt)
        weighted_detail.write_number(r, 1, float(item["productBasicSysNo"]), integer_fmt)
        weighted_detail.write(r, 2, item["selected_source"], external_fmt)
        if item["UnitCost"] is not None:
            weighted_detail.write_number(r, 3, float(item["UnitCost"]), external_cost_fmt)
        weighted_detail.write_number(r, 4, item["selected_record_count"], integer_fmt)
        weighted_detail.write_number(r, 5, item["selected_website_count"], integer_fmt)
        weighted_detail.write(r, 6, item["selected_product_sysnos"], external_fmt)
        weighted_detail.write(r, 7, item["selected_websites"], external_fmt)
        if item["unit_cost_min"] is not None:
            weighted_detail.write_number(r, 8, float(item["unit_cost_min"]), external_cost_fmt)
            weighted_detail.write_number(r, 9, float(item["unit_cost_max"]), external_cost_fmt)
        weighted_detail.write(r, 10, "UnitCost升序；奇数条取中间值，偶数条取中间两条算术平均；0值参与", external_fmt)

    summary.hide_gridlines(2)
    summary.set_column("A:A", 38)
    summary.set_column("B:B", 105)
    summary.write(0, 0, "来福商品池成本及二次定价", title_fmt)
    summary.write(1, 0, "底表", section_fmt)
    summary.write(1, 1, INPUT.name)
    summary.write(2, 0, "成本数据环境", section_fmt)
    summary.write(2, 1, "online")
    summary.write(3, 0, "生成时间", section_fmt)
    summary.write(3, 1, generated_at)
    summary.write(5, 0, "结果指标", section_fmt)
    summary.write(5, 1, "商品数", section_fmt)
    max_row = row_count + 1
    weighted_letter = get_column_letter(col[OLD_WEIGHTED_HEADER] + 1)
    status_letter = get_column_letter(col["定价状态"] + 1)
    delivery_letter = get_column_letter(col["是否产地直发"] + 1)
    cost_type_letter = get_column_letter(col["成本取值口径"] + 1)
    metrics = [
        ("商品池商品总数", f"=COUNTA('二次定价'!$A$2:$A${max_row})", row_count),
        ("有本期末次采购记录", f"=COUNTA('末次采购明细'!$G$2:$G${len(latest) + 1})", stats["有本期末次采购记录"] if stats["有本期末次采购记录"] else len(latest)),
        ("无本期末次采购记录", "=B7-B8", row_count - len(latest)),
        ("有渠道加权成本", f"=COUNT('二次定价'!${weighted_letter}$2:${weighted_letter}${max_row})", stats["有Product_Price记录"]),
        ("无渠道加权成本", f"=COUNTBLANK('二次定价'!${weighted_letter}$2:${weighted_letter}${max_row})", stats["无Product_Price记录"]),
        ("已定价", f"=COUNTIF('二次定价'!${status_letter}$2:${status_letter}${max_row},\"已定价\")", stats["已定价"]),
        ("生鲜／水果本期无采购记录", f"=COUNTIF('二次定价'!${status_letter}$2:${status_letter}${max_row},\"本期无采购记录\")", stats["本期无采购记录"]),
        ("成本缺失", f"=COUNTIF('二次定价'!${status_letter}$2:${status_letter}${max_row},\"成本缺失\")", stats["成本缺失"]),
        ("云仓／代发商品", f"=COUNTIF('二次定价'!${delivery_letter}$2:${delivery_letter}${max_row},\"代发\")+COUNTIF('二次定价'!${delivery_letter}$2:${delivery_letter}${max_row},\"云仓\")", stats["云仓／代发商品"]),
        ("仓发商品", f"=COUNTIF('二次定价'!${delivery_letter}$2:${delivery_letter}${max_row},\"仓发\")", stats["仓发商品"]),
        ("末次采购成本口径商品", f"=COUNTIF('二次定价'!${cost_type_letter}$2:${cost_type_letter}${max_row},\"{PURCHASE}\")", stats[PURCHASE]),
        ("渠道中位数加权成本口径商品", f"=COUNTIF('二次定价'!${cost_type_letter}$2:${cost_type_letter}${max_row},\"{WEIGHTED}\")", stats[WEIGHTED]),
    ]
    for r, (label, formula, cached_value) in enumerate(metrics, start=6):
        summary.write(r, 0, label)
        summary.write_formula(r, 1, formula, integer_fmt, cached_value)
    summary.write(19, 0, "定价规则", section_fmt)
    summary.write(19, 1, "口径说明", section_fmt)
    rules = [
        ("云仓", "原表“代发”按云仓处理：采用成本 ×（1＋2%）。"),
        ("仓发", "净重≤1000g：采用成本 ×（1＋5%）；净重>1000g：采用成本 ×（1＋8%）。"),
        ("成本选择", "水果、蔬菜、肉禽蛋品、水产海鲜、熟食面点使用末次采购成本；其他类目使用Product_Price.UnitCost渠道中位数。"),
        ("取价优先级", "按ProductBasicSysNo关联有效渠道商品：优先分销(ChannelSysNo=5且SaleChannelSysNo=5)；无分销价格记录取C网(SaleChannelSysNo=1)；再无取全部排除集采(SaleChannelSysNo<>2)。"),
        ("中位数", "同一优先级下多站点多条UnitCost升序取中位数；偶数条取中间两条算术平均，0值参与计算；不按Status上下架状态过滤。"),
        ("末次采购成本", "来自newretail_oms本期采购收货／退货明细，按采购时间和回调系统编号倒序取每商品最新一条，不使用默认采购价替代。"),
        ("缺失成本", "无有效渠道价格记录、UnitCost为空或中位数≤0时不计算二次定价，定价状态标为“成本缺失”。"),
        ("价格精度", "二次定价四舍五入保留2位小数；主表计算列保留Excel公式。"),
    ]
    for r, (label, text) in enumerate(rules, start=20):
        summary.write(r, 0, label)
        summary.write(r, 1, text)
    summary.write(29, 0, "工作表", section_fmt)
    summary.write(29, 1, "内容", section_fmt)
    sheet_index = [
        ("二次定价", "沿用当前31列底表；加权成本列改用渠道UnitCost中位数，重算采用成本、二次定价和定价状态。"),
        ("末次采购明细", f"{len(latest):,}条末次采购记录，含时间、单号、仓库、类型、数量、成本和税率。"),
        ("加权成本取价明细", f"{len(erp):,}个商品的取价渠道、中位数、记录数、站点数、渠道商品SysNo及成本区间。"),
        ("定价参数", "加价率、重量阈值、类目成本口径与执行假设。"),
        ("数据来源", "线上接口、数据库查询、底表和字段映射说明。"),
    ]
    for r, (label, text) in enumerate(sheet_index, start=30):
        summary.write(r, 0, label)
        summary.write(r, 1, text)

    sources.hide_gridlines(2)
    sources.set_column("A:A", 42)
    sources.set_column("B:B", 105)
    sources.set_column("C:C", 110)
    sources.write_row(0, 0, ["Source Name", "Source URL", "用途与口径"], section_fmt)
    output_source_rows = []
    for row in small_formula["数据来源"][1:]:
        values = list(row)
        for c, value in enumerate(values):
            if isinstance(value, str):
                value = value.replace("costPrice为商品主数据单位成本，用作非生鲜／水果类目的加权成本。", "商品主数据costPrice仅保留作对照，不参与本次加权成本定价。")
                value = value.replace("costPrice及来福池成本", "来福池成本")
                values[c] = value
        if any("ProductService.cs" in str(value) for value in values if value is not None):
            values[2] = "用于核对分销、C网、集采渠道编码和IsDelete有效商品口径；本次按新要求改为渠道优先级与中位数。"
        output_source_rows.append(values)
    output_source_rows.append([
        "ERP商品多渠道映射 SOA（online）",
        "https://vm-online.benlai.com/erp-product-webapi/ProLhotse/GetListByBasic",
        "按ProductBasicSysNo批量获取ProductSysNo、WebSiteSysNo、ChannelSysNo、SaleChannelSysNo、Status、IsDelete；仅排除IsDelete=true。",
    ])
    output_source_rows.append([
        "Archery线上只读查询（benlai.dbo.Product_Price）",
        "https://dbaudit.benlai.com/sqlquery/",
        "实例[online]P0-sqlserver-master-104-138；按全部有效渠道商品ProductSysNo采集UnitCost，无任何时间周期或LastUpdateTime过滤；再按分销→C网→非集采全渠道优先级取中位数。",
    ])
    output_source_rows.append(["渠道中位数成本调整时间", generated_at, "UnitCost为含税单位成本；NoTaxUnitCost不参与本次定价；0值参与中位数。"])
    for r, row in enumerate(output_source_rows, start=1):
        for c, value in enumerate(row):
            if value is not None:
                sources.write_string(r, c, str(value))

    source_formula.close()
    source_values.close()
    workbook.close()
    result = {
        "output": str(OUTPUT),
        "rows": row_count,
        "columns": len(headers),
        "product_price_records": len(erp),
        "products_without_product_price": stats["无Product_Price记录"],
        "latest_purchase_records": len(latest),
        "status": {key: stats[key] for key in ("已定价", "本期无采购记录", "成本缺失", "发货方式或重量缺失")},
        "cost_types": {PURCHASE: stats[PURCHASE], WEIGHTED: stats[WEIGHTED]},
        "file_size_bytes": OUTPUT.stat().st_size,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
