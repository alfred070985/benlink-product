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
INPUT = ROOT / "来福商品池_成本及二次定价_2026-10-30.xlsx"
LATEST = ROOT / "latest_purchase_costs_2026-10-01_2026-10-30.jsonl"
PROGRESS = ROOT / "latest_purchase_costs_progress.json"
COUNT_JSON = ROOT / "purchase_callbacks_count.json"
OUTPUT = ROOT / "来福商品池_成本及二次定价_末次采购成本更新_2026-10-30.xlsx"
SPECIAL = {"水果", "蔬菜", "肉禽蛋品", "水产海鲜", "熟食面点"}
PURCHASE = "末次采购成本（采购明细）"
WEIGHTED = "加权成本（costPrice）"
OLD_PURCHASE = "采购成本（purchasePrice）"


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


def load_latest():
    rows = {}
    with LATEST.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                item = json.loads(line)
                pid = item["product_id"]
                if pid in rows:
                    raise RuntimeError(f"Duplicate latest purchase product: {pid}")
                rows[pid] = item
    return rows


def main():
    progress = json.loads(PROGRESS.read_text(encoding="utf-8"))
    if not progress.get("coverage_complete"):
        raise RuntimeError("Purchase collection is not complete")
    latest = load_latest()
    raw_detail_count = json.loads(COUNT_JSON.read_text(encoding="utf-8"))["data"]["rows"][0][0]
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    source_formula = load_workbook(INPUT, read_only=True, data_only=False)
    source_values = load_workbook(INPUT, read_only=True, data_only=True)
    small_formula = {}
    small_values = {}
    for name in ("汇总", "定价参数", "数据来源"):
        small_formula[name] = list(source_formula[name].iter_rows(values_only=True))
        small_values[name] = list(source_values[name].iter_rows(values_only=True))
    sf = source_formula["二次定价"]
    sv = source_values["二次定价"]
    formula_iter = sf.iter_rows(values_only=True)
    value_iter = sv.iter_rows(values_only=True)
    headers = list(next(formula_iter))
    cached_headers = list(next(value_iter))
    if headers != cached_headers or len(headers) != 31:
        raise RuntimeError("Unexpected source workbook structure")
    col = {name: idx for idx, name in enumerate(headers)}
    required = ["商品编号", "商品大类", "是否产地直发", "净重(g)", "加权成本costPrice(元)",
                "采购成本purchasePrice(元)", "成本取值口径", "采用成本(元)", "重量分档", "加价率", "二次定价(元)", "定价状态"]
    if any(name not in col for name in required):
        raise RuntimeError("Required source columns are missing")

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

    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, sf.max_row - 1, len(headers) - 1)
    ws.set_row(0, 30)
    for c, name in enumerate(headers):
        ws.write(0, c, "末次采购成本(元)" if name == "采购成本purchasePrice(元)" else name, header_fmt)
    widths = {
        "商品编号": 16, "商品名称": 42, "商品大类": 16, "商品中类": 18, "商品小类": 20,
        "商品池售价": 14, "价格下限": 14, "价格上限": 14, "是否产地直发": 16,
        "净重(g)": 13, "成本取值口径": 28, "采用成本(元)": 16, "重量分档": 16,
        "加价率": 12, "二次定价(元)": 16, "定价状态": 24, "采购成本purchasePrice(元)": 22,
    }
    for name, idx in col.items():
        ws.set_column(idx, idx, widths.get(name, 15))

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
        item = latest.get(pid)
        purchase = number(item["buy_price"]) if item is not None else None
        old_purchase = number(cached[col["采购成本purchasePrice(元)"]])
        values[col["采购成本purchasePrice(元)"]] = purchase
        cached[col["采购成本purchasePrice(元)"]] = purchase
        if item is not None:
            stats["有本期末次采购记录"] += 1
        else:
            stats["无本期末次采购记录"] += 1
        if old_purchase is None and purchase is not None:
            stats["成本新增"] += 1
        elif old_purchase is not None and purchase is None:
            stats["成本清空"] += 1
        elif old_purchase is not None and purchase is not None and abs(old_purchase - purchase) > 1e-9:
            stats["成本变更"] += 1
        else:
            stats["成本不变"] += 1

        category = values[col["商品大类"]]
        delivery = values[col["是否产地直发"]]
        weight = number(cached[col["净重(g)"]])
        weighted = number(cached[col["加权成本costPrice(元)"]])
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

        def ref(name):
            return f"{get_column_letter(col[name] + 1)}{excel_row}"

        cost_type_ref = ref("成本取值口径")
        purchase_ref = ref("采购成本purchasePrice(元)")
        weighted_ref = ref("加权成本costPrice(元)")
        selected_ref = ref("采用成本(元)")
        rate_ref = ref("加价率")
        formulas = {
            "成本取值口径": values[col["成本取值口径"]],
            "采用成本(元)": f'=IF({cost_type_ref}="{PURCHASE}",{purchase_ref},{weighted_ref})',
            "重量分档": values[col["重量分档"]],
            "加价率": values[col["加价率"]],
            "二次定价(元)": values[col["二次定价(元)"]],
            "定价状态": f'=IF(AND({cost_type_ref}="{PURCHASE}",{purchase_ref}=""),"本期无采购记录",IF(OR({selected_ref}="",{selected_ref}<=0),"成本缺失",IF({rate_ref}="","发货方式或重量缺失","已定价")))',
        }
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
        for idx, name in enumerate(headers):
            value = values[idx]
            if name in formulas:
                cached_value = formula_cache[name]
                ws.write_formula(out_row, idx, formulas[name], formula_formats[name], "" if cached_value is None else cached_value)
            elif value is not None:
                if name == "采购成本purchasePrice(元)":
                    ws.write_number(out_row, idx, value, external_cost_fmt)
                elif 17 <= idx <= 24:
                    ws.write(out_row, idx, safe(value), external_fmt)
                else:
                    ws.write(out_row, idx, safe(value))
    extra_formula_rows = sum(1 for _ in formula_iter)
    extra_value_rows = sum(1 for _ in value_iter)
    if extra_formula_rows or extra_value_rows or row_count != 43519:
        raise RuntimeError("Source row count mismatch")
    if pool_ids - set(latest) != set() and len(pool_ids) != row_count:
        raise RuntimeError("Product identity mismatch")
    if stats["有本期末次采购记录"] != len(latest):
        raise RuntimeError("Latest purchase records do not all belong to the product pool")

    # 参数页：保留用户当前参数和类目表，只替换采购成本来源说明。
    params.hide_gridlines(2)
    params.set_column("A:A", 110)
    params.set_column("B:B", 18)
    params.set_column("D:D", 24)
    params.set_column("E:E", 32)
    for r, row in enumerate(small_formula["定价参数"]):
        for c, original in enumerate(row):
            value = original
            if value == OLD_PURCHASE:
                value = PURCHASE
            if isinstance(value, str) and value.startswith("purchasePrice在商品主数据"):
                value = "末次采购成本来自newretail_oms采购收货／退货明细，日期范围[2026-10-01, 2026-10-30)，按inventory_time、callback_sys_no倒序取每商品最新一条；不再使用商品主数据purchasePrice。"
            if isinstance(value, str) and value.startswith("采用成本为空或≤0时"):
                value = "采用成本为空或≤0时不计算二次定价；生鲜／水果类目无本期采购记录时，定价状态标为“本期无采购记录”。"
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
    params.set_row(33, 40)

    # 原始末次采购明细。
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

    # 汇总页公式及缓存。
    summary.hide_gridlines(2)
    summary.set_column("A:A", 34)
    summary.set_column("B:B", 100)
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
    metrics = [
        ("商品池商品总数", f"=COUNTA('二次定价'!$A$2:$A${max_row})", row_count),
        ("有本期末次采购记录", f"=COUNTA('末次采购明细'!$G$2:$G${len(latest) + 1})", stats["有本期末次采购记录"]),
        ("无本期末次采购记录", "=B7-B8", stats["无本期末次采购记录"]),
        ("已定价", f"=COUNTIF('二次定价'!$AE$2:$AE${max_row},\"已定价\")", stats["已定价"]),
        ("生鲜／水果本期无采购记录", f"=COUNTIF('二次定价'!$AE$2:$AE${max_row},\"本期无采购记录\")", stats["本期无采购记录"]),
        ("成本缺失", f"=COUNTIF('二次定价'!$AE$2:$AE${max_row},\"成本缺失\")", stats["成本缺失"]),
        ("云仓／代发商品", f"=COUNTIF('二次定价'!$M$2:$M${max_row},\"代发\")+COUNTIF('二次定价'!$M$2:$M${max_row},\"云仓\")", stats["云仓／代发商品"]),
        ("仓发商品", f"=COUNTIF('二次定价'!$M$2:$M${max_row},\"仓发\")", stats["仓发商品"]),
        ("末次采购成本口径商品", f"=COUNTIF('二次定价'!$Z$2:$Z${max_row},\"{PURCHASE}\")", stats[PURCHASE]),
        ("加权成本口径商品", f"=COUNTIF('二次定价'!$Z$2:$Z${max_row},\"{WEIGHTED}\")", stats[WEIGHTED]),
    ]
    for r, (label, formula, cached_value) in enumerate(metrics, start=6):
        summary.write(r, 0, label)
        summary.write_formula(r, 1, formula, integer_fmt, cached_value)
    summary.write(17, 0, "定价规则", section_fmt)
    summary.write(17, 1, "口径说明", section_fmt)
    rules = [
        ("云仓", "原表“代发”按云仓处理：采用成本 ×（1＋2%）。"),
        ("仓发", "净重≤1000g：采用成本 ×（1＋5%）；净重>1000g：采用成本 ×（1＋8%）。"),
        ("成本选择", "水果、蔬菜、肉禽蛋品、水产海鲜、熟食面点使用末次采购成本；其他类目使用costPrice加权成本。"),
        ("末次采购成本", "来自newretail_oms本期采购收货／退货明细，按采购时间和回调系统编号倒序取每商品最新一条，不使用默认采购价替代。"),
        ("无本期采购记录", "末次采购成本留空；生鲜／水果类目标记“本期无采购记录”，其他类目仍按加权成本定价。"),
        ("缺失成本", "有记录但成本为空或≤0时不计算二次定价，定价状态标为“成本缺失”。"),
        ("价格精度", "二次定价四舍五入保留2位小数；主表计算列保留Excel公式。"),
    ]
    for r, (label, text) in enumerate(rules, start=18):
        summary.write(r, 0, label)
        summary.write(r, 1, text)
    summary.write(26, 0, "工作表", section_fmt)
    summary.write(26, 1, "内容", section_fmt)
    sheet_index = [
        ("二次定价", "用户当前31列底表；仅更新末次采购成本及相关公式结果，未恢复已删除列。"),
        ("末次采购明细", f"{len(latest):,}条末次采购记录，含时间、单号、仓库、类型、数量、成本和税率。"),
        ("定价参数", "加价率、重量阈值、类目成本口径与执行假设。"),
        ("数据来源", "线上接口、数据库查询、底表和字段映射说明。"),
    ]
    for r, (label, text) in enumerate(sheet_index, start=27):
        summary.write(r, 0, label)
        summary.write(r, 1, text)

    sources.hide_gridlines(2)
    sources.set_column("A:A", 40)
    sources.set_column("B:B", 105)
    sources.set_column("C:C", 105)
    sources.write_row(0, 0, ["Source Name", "Source URL", "用途与口径"], section_fmt)
    source_rows = [
        ["本来生活商品主数据 SOA（online）", "https://vm-online.benlai.com/business-product-servicehost/ProductApi/GetProductListByIds", "读取商品系统编号、主商品编号、净重、发货方式、costPrice及来福池成本；purchasePrice不再作为末次采购成本来源。"],
        [INPUT.name, "本地底表", "以用户删减后的31列版本为唯一底表，保留全部商品和未指定修改的列。"],
        ["线上代码：ProductService.cs", "benlai.business.product／master", "costPrice为商品主数据单位成本，用作非生鲜／水果类目的加权成本。"],
        ["Archery线上只读查询（newretail_oms）", "https://dbaudit.benlai.com/sqlquery/", "查询callback_wmsx_order、callback_wmsx_order_detail及benlai_product.product_basic_info；日期[2026-10-01, 2026-10-30)，customer_code=BLSH，保留指定采购收货／退货类型。"],
        ["本期采购明细总量", "https://dbaudit.benlai.com/sqlquery/", f"按用户给定FROM/WHERE口径COUNT(*)共{raw_detail_count:,}行；商品池按商品编号区间分批采集每商品最新一条。"],
        ["末次采购排序口径", "https://dbaudit.benlai.com/sqlquery/", "ROW_NUMBER按商品分区，inventory_time DESC、callback_sys_no DESC，再以单号、商品系统编号、数量、单位成本确定顺序；取第1条。"],
        ["数据提取及文件生成时间", generated_at, f"商品池{row_count:,}个；有末次采购记录{len(latest):,}个；无本期采购记录{row_count - len(latest):,}个。"],
    ]
    for r, row in enumerate(source_rows, start=1):
        for c, value in enumerate(row):
            sources.write_string(r, c, str(value))

    source_formula.close()
    source_values.close()
    workbook.close()
    result = {
        "output": str(OUTPUT),
        "rows": row_count,
        "columns": len(headers),
        "purchase_records": len(latest),
        "no_purchase_records": stats["无本期无采购记录"] if "无本期无采购记录" in stats else stats["无本期末次采购记录"],
        "raw_detail_rows": raw_detail_count,
        "status": {key: stats[key] for key in ("已定价", "本期无采购记录", "成本缺失", "发货方式或重量缺失")},
        "cost_changes": {key: stats[key] for key in ("成本新增", "成本清空", "成本变更", "成本不变")},
        "file_size_bytes": OUTPUT.stat().st_size,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
