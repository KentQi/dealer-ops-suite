#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""月结买断独立对账引擎：签收净额结货款，乙方动销只计算贴花。

通过 RECON_CFG 注入与铺货代销相同的客户／期间／主档上下文，另需：
  contract.method=月结买断；contract.buyout_date 与 payment_due_date 分别登记；
  sources.receipts / returns / sales 指向签收、退货、乙方动销三份 Excel。
默认只生成结果，不回写主档；_回写.json 交 writeback_ledger.py 复核后使用。
"""

import json
import os
import sys
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from schema_common import (  # noqa: E402
    ColumnAmbiguous, ColumnMissing, locate_column, normalize_barcode, sha256_file,
)
from settlement_rules import (  # noqa: E402
    RateBook, RuleViolation, calculate_monthly_buyout, jsonable, normalize_business_type,
)

NAVY, BLUE, WHITE, GREY, RED = "1F3864", "D9E2F3", "FFFFFF", "F2F2F2", "C00000"
THIN = Side(style="thin", color="BFBFBF")


def fail(message):
    raise SystemExit("[中止] " + str(message))


_REAL_OUT = sys.stdout
sys.stdout = sys.stderr          # 与铺货引擎一致：过程输出走 stderr；成功交付行走 emit 到 stdout


def emit(msg=""):
    _REAL_OUT.write(str(msg) + "\n")
    _REAL_OUT.flush()


def load_config():
    try:
        cfg = json.loads(os.environ.get("RECON_CFG", ""))
    except (ValueError, TypeError) as exc:
        fail("RECON_CFG 必须是合法 JSON：" + str(exc))
    if not isinstance(cfg, dict):
        fail("RECON_CFG 顶层必须是对象")
    return cfg


def bar(value):
    """条码归一化：委托共享内核（None→""；整值 float→int 串；".0" 尾巴剥除）。"""
    return normalize_barcode(value)


def column(headers, aliases, required=True):
    """列定位：委托共享内核 locate_column（精确优先→子串回退→泛化词守卫→多列阻断）。

    fail() 文案与旧实现逐字一致（sheet_name 留空使内核前缀为「源表」），由探针逐条比对钉死。
    """
    try:
        return locate_column(headers, aliases, required=required)
    except (ColumnAmbiguous, ColumnMissing) as exc:
        fail(str(exc))


def read_rows(path, date_aliases):
    if not path or not os.path.isfile(path):
        fail("源文件不存在：" + str(path))
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        candidates = [ws for ws in wb.worksheets if ws.max_row > 0]
        if len(candidates) != 1:
            fail("每份月结源文件必须只有一张工作表，避免选错：" + path)
        ws = candidates[0]
        matrix = [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    header_row = None
    for index, row in enumerate(matrix[:6]):
        if any(any(alias in str(value or "") for alias in date_aliases) for value in row):
            header_row = index
            break
    if header_row is None:
        fail("未找到业务日期表头：" + path)
    return matrix[header_row], matrix[header_row + 1:]


def rows_of(path, role, bool_reader=False):
    if role == "receipts":
        headers, matrix = read_rows(path, ("签收日期", "签收时间"))
        positions = {
            "id": column(headers, ("流水号", "签收单号")),
            "date": column(headers, ("签收日期", "签收时间")),
            "barcode": column(headers, ("商品条码", "条码")),
            "quantity": column(headers, ("签收数量", "数量")),
        }
    elif role == "returns":
        headers, matrix = read_rows(path, ("退货日期", "退货时间"))
        positions = {
            "id": column(headers, ("流水号", "退货单号")),
            "date": column(headers, ("退货日期", "退货时间")),
            "barcode": column(headers, ("商品条码", "条码")),
            "quantity": column(headers, ("退货数量", "数量")),
            "original_receipt_id": column(headers, ("原签收流水号", "原签收单号"), False),
            "original_sale_id": column(headers, ("原销售流水号", "原销售单号"), False),
            "original_unit_price": column(headers, ("原签收单价", "原销售单价", "原结算价"), False),
            "original_decal_charged": column(headers, ("原销售已计提贴花", "原销售贴花已计提"), False),
            "original_decal_standard": column(headers, ("原销售贴花标准", "原贴花标准"), False),
        }
    else:
        headers, matrix = read_rows(path, ("动销日期", "销售日期", "销售时间"))
        positions = {
            "id": column(headers, ("流水号", "销售单号")),
            "date": column(headers, ("动销日期", "销售日期", "销售时间")),
            "barcode": column(headers, ("商品条码", "条码")),
            "quantity": column(headers, ("动销数量", "销售数量", "数量")),
        }
    output = []
    for row_number, values in enumerate(matrix, 2):
        if not values or all(value in (None, "") for value in values):
            continue
        def get(key):
            index = positions.get(key)
            return values[index] if index is not None and index < len(values) else None
        if get("id") in (None, ""):
            fail(f"{path} 第 {row_number} 行缺少流水号")
        item = {"id": str(get("id")).strip(), "barcode": bar(get("barcode")),
                "date": get("date"), "quantity": get("quantity")}
        for key in ("original_receipt_id", "original_sale_id", "original_unit_price",
                    "original_decal_charged", "original_decal_standard"):
            if key in positions and get(key) not in (None, ""):
                item[key] = get(key)
        if "original_decal_charged" in item:
            item["original_decal_charged"] = bool_reader(item["original_decal_charged"], path, row_number)
        output.append(item)
    return output


def parse_bool(value, path, row_number):
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in ("1", "true", "是", "已计提", "有"):
        return True
    if text in ("0", "false", "否", "未计提", "无"):
        return False
    fail(f"{path} 第 {row_number} 行原销售是否计提贴花只能填是／否")


def master_sheet_name(mapping, name):
    """主档工作表名：_数据映射.json 的 master.sheets 声明了改名时用实际名（与铺货引擎同一口径）。"""
    if not isinstance(mapping, dict):
        return name
    master_map = mapping.get("master")
    if not isinstance(master_map, dict):
        return name
    declared = master_map.get("sheets")
    if not isinstance(declared, dict):
        return name
    return {str(canonical): str(actual) for actual, canonical in declared.items()}.get(name, name)


def load_rate_book(master, sheet_name, value_aliases, required, kind):
    from settlement_rules import rate_value_column

    if sheet_name not in master.sheetnames:
        if required:
            fail("主档缺少「" + sheet_name + "」表")
        return None
    ws = master[sheet_name]
    headers = [cell.value for cell in ws[1]]
    ci_bar = column(headers, ("商品条码", "条码"))
    ci_value = rate_value_column(headers, value_aliases, sheet_name)
    ci_start = column(headers, ("生效起",))
    ci_end = column(headers, ("生效止",), False)
    ci_source = column(headers, ("依据文件",), False)
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not values or values[ci_bar] in (None, ""):
            continue
        rows.append({"barcode": values[ci_bar], "value": values[ci_value], "start": values[ci_start],
                     "end": values[ci_end] if ci_end is not None else None,
                     "source": values[ci_source] if ci_source is not None else ""})
    if required and not rows:
        fail("主档「" + sheet_name + "」没有有效价格记录")
    return RateBook(rows, kind=kind)


def workbook_sha256(path):
    """工作簿 sha256：委托共享内核（1 MiB 分块流式摘要）。"""
    return sha256_file(path)


def json_write(path, value):
    temp = path + ".tmp"
    with open(temp, "w", encoding="utf-8") as fh:
        json.dump(value, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    os.replace(temp, path)


def table(ws, start, headers, rows, money_columns=(), widths=None):
    for col, header in enumerate(headers, 1):
        cell = ws.cell(start, col, header)
        cell.font = Font(name="微软雅黑", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    ws.row_dimensions[start].height = 30
    for row_index, values in enumerate(rows, start + 1):
        for col, value in enumerate(values, 1):
            cell = ws.cell(row_index, col, value)
            cell.font = Font(name="微软雅黑", size=9)
            cell.fill = PatternFill("solid", fgColor=WHITE if row_index % 2 else GREY)
            cell.alignment = Alignment(horizontal="right" if col in money_columns else "center",
                                       vertical="center", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            if col in money_columns:
                cell.number_format = "#,##0.00"
        ws.row_dimensions[row_index].height = 23
    for index, width in enumerate(widths or [16] * len(headers), 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A" + str(start + 1)
    return start + len(rows) + 1


def title(ws, text, subtitle, columns):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=columns)
    cell = ws.cell(1, 1, text)
    cell.font = Font(name="微软雅黑", size=20, bold=True, color=NAVY)
    cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 38
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=columns)
    cell = ws.cell(2, 1, subtitle)
    cell.font = Font(name="微软雅黑", size=9, color="595959")
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[2].height = 28
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def main():
    cfg = load_config()
    period = cfg.get("period") or {}
    contract = cfg.get("contract") or {}
    sources = cfg.get("sources") or {}
    base = cfg.get("base_dir") or ""
    sources = {role: os.path.abspath(value if os.path.isabs(str(value)) else os.path.join(base, str(value)))
               for role, value in sources.items() if value}
    required = {
        "甲方名称": (cfg.get("party_a") or {}).get("name"),
        "乙方名称": (cfg.get("party_b") or {}).get("name"),
        "期间起": period.get("start"), "期间止": period.get("end"),
        "主档路径": cfg.get("master_path"), "签收文件": sources.get("receipts"),
        "退货文件": sources.get("returns"), "乙方动销文件": sources.get("sales"),
        "买断日": contract.get("buyout_date") or period.get("buyout_date"),
        "付款截止日": contract.get("payment_due_date") or period.get("payment_due_date"),
    }
    missing = [label for label, value in required.items() if not value]
    if missing:
        fail("缺少必要信息：" + "／".join(missing) + "；买断日与付款截止日不得默认或合并")
    try:
        business_type = normalize_business_type(cfg.get("business_type") or contract.get("method"))
        if business_type != "monthly_buyout":
            fail("本引擎仅服务月结买断；铺货代销请用 recon_engine.py")
        receipts = rows_of(sources["receipts"], "receipts")
        returns = rows_of(sources["returns"], "returns", parse_bool)
        sales = rows_of(sources["sales"], "sales")
        mapping = cfg.get("mapping")
        if mapping is not None and (not isinstance(mapping, dict)
                                    or mapping.get("schema") not in (None, "dealer-ops/data-mapping@1")):
            fail("RECON_CFG.mapping 不支持：" + repr(mapping.get("schema") if isinstance(mapping, dict) else mapping))
        master = openpyxl.load_workbook(cfg["master_path"], data_only=True)
        prices = load_rate_book(master, master_sheet_name(mapping, "价格"),
                                ("供货价", "结算价", "合同价"), True, "合同价")
        decals = load_rate_book(master, master_sheet_name(mapping, "贴花"),
                                ("贴花标准",), False, "贴花标准")
        names, attributes = {}, {}
        sku_sheet = master_sheet_name(mapping, "SKU")
        if sku_sheet not in master.sheetnames:
            fail("主档缺少 SKU 表；不得默认全部按结算类处理")
        ws = master[sku_sheet]
        headers = [cell.value for cell in ws[1]]
        ci_bar = column(headers, ("商品条码", "条码"))
        ci_name = column(headers, ("标准品名", "商品名称", "品名"), False)
        ci_attr = column(headers, ("结算属性",))
        for values in ws.iter_rows(min_row=2, values_only=True):
            if values and values[ci_bar] not in (None, ""):
                barcode = bar(values[ci_bar])
                attribute = str(values[ci_attr] or "").strip()
                if barcode in attributes and attributes[barcode] != attribute:
                    fail("同一 SKU 的结算属性登记不一致：" + barcode)
                attributes[barcode] = attribute
                names[barcode] = str(values[ci_name] or "") if ci_name is not None else ""
        master.close()
        bars = {line["barcode"] for line in receipts + returns + sales}
        unknown = sorted(barcode for barcode in bars if attributes.get(barcode) not in ("结算类", "非结算类"))
        if unknown:
            fail("以下 SKU 未明确登记结算类／非结算类：" + "、".join(unknown))
        settled = {barcode for barcode in bars if attributes[barcode] == "结算类"}
        ignored = sorted(bars - settled)
        receipts = [line for line in receipts if line["barcode"] in settled]
        returns = [line for line in returns if line["barcode"] in settled]
        sales = [line for line in sales if line["barcode"] in settled]
        bars = settled
        for barcode in sorted(bars):
            prices.validate_period(barcode, period["start"], period["end"])
        result = calculate_monthly_buyout(
            period_start=period["start"], period_end=period["end"],
            receipts=receipts, returns=returns, sales=sales, prices=prices, decals=decals,
            original_sales=cfg.get("original_sales") or [],
            original_receipts=cfg.get("original_receipts") or [],
            buyout_date=contract.get("buyout_date") or period.get("buyout_date"),
            payment_due_date=contract.get("payment_due_date") or period.get("payment_due_date"),
        )
    except RuleViolation as exc:
        fail(f"结算规则阻断[{exc.code}] {exc.message}")
    finally:
        try:
            master.close()
        except (NameError, AttributeError):
            pass

    out_path = cfg.get("out_file") or os.path.join(
        base, period["end"][:4] + period["end"][5:7] + "月结买断对账.xlsx")
    if os.path.exists(out_path) and not cfg.get("overwrite"):
        fail("输出文件已存在；请改名或显式设置 overwrite=true：" + out_path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    party_text = f"{(cfg.get('party_a') or {}).get('name')} × {(cfg.get('party_b') or {}).get('name')}"
    subtitle = f"{party_text}｜{period['start']} 至 {period['end']}｜买断日 {result['buyout_date']}｜付款截止日 {result['payment_due_date']}"

    ws = wb.create_sheet("00-对账说明")
    title(ws, "月结买断对账说明", subtitle, 4)
    notes = [
        ("业务类型", "月结买断；直接买断不进入本对账引擎"),
        ("数量规则", result["quantity_rule"]),
        ("价格规则", result["price_rule"]),
        ("贴花规则", result["decal_pricing_rule"]),
        ("本期签收／退货／买断数量", f"{result['receipt_qty']}／{result['return_qty']}／{result['buyout_qty']}"),
        ("签收货款／退货冲减／买断货款", f"{result['gross_receipt_amount']}／{result['return_amount']}／{result['buyout_amount']}"),
        ("乙方动销数量（只算贴花）", str(result["sales_qty_for_decal_only"])),
        ("贴花金额（独立于货款）", str(result["decal_amount"])),
        ("日期控制", "买断日与付款截止日分别记录；付款日按合同原文，不自行顺延或提前"),
    ]
    table(ws, 4, ["项目", "内容", "", ""], [[key, value, None, None] for key, value in notes], widths=[24, 30, 30, 30])
    for row in range(5, 5 + len(notes)):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)

    receipt_by_bar = defaultdict(lambda: {"qty": 0, "amount": 0.0})
    return_by_bar = defaultdict(lambda: {"qty": 0, "amount": 0.0, "decal": 0.0})
    sale_by_bar = defaultdict(lambda: {"qty": 0, "decal": 0.0})
    for line in result["priced_receipts"]:
        receipt_by_bar[line["barcode"]]["qty"] += float(line["quantity"])
        receipt_by_bar[line["barcode"]]["amount"] += float(line["amount"])
    for line in result["priced_returns"]:
        return_by_bar[line["barcode"]]["qty"] -= float(line["quantity"])
        return_by_bar[line["barcode"]]["amount"] += float(line["amount"])
        return_by_bar[line["barcode"]]["decal"] += float(line["decal_amount"])
    for line in result["priced_sales_for_decal"]:
        sale_by_bar[line["barcode"]]["qty"] += float(line["quantity"])
        sale_by_bar[line["barcode"]]["decal"] += float(line["decal_amount"])

    ws = wb.create_sheet("01-月结对账单")
    title(ws, "月结买断对账单", subtitle, 10)
    summary_rows = []
    for barcode in sorted(bars):
        summary_rows.append([barcode, names.get(barcode) or "（主档未登记品名）",
                             receipt_by_bar[barcode]["qty"], return_by_bar[barcode]["qty"],
                             float(result["buyout_qty"]) if len(bars) == 1 else receipt_by_bar[barcode]["qty"] - return_by_bar[barcode]["qty"],
                             sale_by_bar[barcode]["qty"], receipt_by_bar[barcode]["amount"],
                             return_by_bar[barcode]["amount"],
                             receipt_by_bar[barcode]["amount"] + return_by_bar[barcode]["amount"],
                             sale_by_bar[barcode]["decal"] + return_by_bar[barcode]["decal"]])
    summary_rows.append(["合计", "", float(result["receipt_qty"]), float(result["return_qty"]),
                         float(result["buyout_qty"]), float(result["sales_qty_for_decal_only"]),
                         float(result["gross_receipt_amount"]), float(result["return_amount"]),
                         float(result["buyout_amount"]), float(result["decal_amount"])])
    table(ws, 4, ["商品条码", "商品名称", "本期签收", "本期退货", "买断数量", "动销数量\n（仅贴花）",
                  "签收货款", "退货冲减", "买断货款", "贴花金额"], summary_rows,
          money_columns=(7, 8, 9, 10), widths=[16, 26, 11, 11, 11, 14, 14, 14, 14, 14])

    ws = wb.create_sheet("02-签收与退货")
    title(ws, "签收与退货逐笔取价", subtitle, 10)
    audit_rows = []
    for kind, lines in (("签收", result["priced_receipts"]), ("退货", result["priced_returns"])):
        for line in lines:
            audit_rows.append([kind, line["id"], line["date"], line["barcode"], names.get(line["barcode"], ""),
                               float(line["quantity"]), float(line["unit_price"]), float(line["amount"]),
                               line.get("original_receipt_id") or "—", line["price_evidence"]])
    table(ws, 4, ["类型", "流水号", "业务日期", "商品条码", "商品名称", "数量", "合同价／原签收价",
                  "货款金额", "原签收流水", "取价依据"], audit_rows, money_columns=(7, 8),
          widths=[8, 16, 13, 16, 24, 10, 16, 14, 16, 32])

    ws = wb.create_sheet("03-乙方动销贴花")
    title(ws, "乙方动销贴花（不减少买断货款）", subtitle, 9)
    decal_rows = []
    for kind, lines in (("动销", result["priced_sales_for_decal"]), ("退货冲减", result["priced_returns"])):
        for line in lines:
            decal_rows.append([kind, line["id"], line["date"], line["barcode"], names.get(line["barcode"], ""),
                               float(line["quantity"]),
                               float(line["decal_standard"]) if line.get("decal_standard") is not None else None,
                               float(line["decal_amount"]), line.get("price_evidence") or "按原销售贴花标准冲减"])
    table(ws, 4, ["类型", "流水号", "动销／退货日期", "商品条码", "商品名称", "数量", "贴花标准",
                  "贴花金额", "依据"], decal_rows, money_columns=(7, 8), widths=[10, 16, 15, 16, 24, 10, 14, 14, 34])

    temp = out_path + ".tmp.xlsx"
    wb.save(temp)
    os.replace(temp, out_path)
    workbook_sha = workbook_sha256(out_path)
    period_key = period["end"][:4] + period["end"][5:7]
    writeback = {
        "period": period_key, "asof": period.get("asof") or period["end"],
        "opening_qty": 0, "opening_amount": 0,
        "settle_qty": float(result["buyout_qty"]), "settle_amount": float(result["buyout_amount"]),
        "diff_qty": 0, "diff_amount": 0,
        "sticker_amount": float(result["decal_amount"]), "sticker_pay": "待定", "status": "待确认",
        "contract_no": contract.get("no") or "", "business_type": "月结买断",
        "statement_no": str(cfg.get("statement_no") or "").strip(),
        "buyout_date": result["buyout_date"], "payment_due_date": result["payment_due_date"],
        "note": f"月结买断；买断日 {result['buyout_date']}；付款截止日 {result['payment_due_date']}；动销只计算贴花。",
    }
    result["ignored_non_settlement_barcodes"] = ignored
    result_path = os.path.join(base, "_结算结果.json")
    writeback_path = os.path.join(base, "_回写.json")
    json_write(writeback_path, writeback)
    json_write(result_path, {
        "schema": "dealer-ops/settlement-result@1", "business_type": business_type,
        "workbook": {"path": os.path.abspath(out_path), "sha256": workbook_sha},
        "rule_result": jsonable(result), "writeback": writeback, "writeback_path": writeback_path,
        "master": {"path": os.path.abspath(cfg["master_path"]),
                   "sha256": workbook_sha256(cfg["master_path"])},
        "sources": {role: {"path": os.path.abspath(sources[role]),
                           "sha256": workbook_sha256(sources[role])}
                    for role in ("receipts", "returns", "sales")},
    })
    emit("SAVED: " + out_path)
    emit("RESULT: " + result_path)
    emit("WRITEBACK: " + writeback_path)


if __name__ == "__main__":
    main()
