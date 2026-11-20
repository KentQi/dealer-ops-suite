#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把已核验的 _结算结果.json 投影为铺货／月结各自的复核底稿。"""
import json
import os
import re
from decimal import Decimal

# 合计列的舍入口径收敛到权威金额规则：Decimal 和先按 HALF_UP 量化到分，再交 _number 渲染，
# 保证「合计 == 逐行舍入后之和」不变式（_number 本体是 JSON 数值渲染归一器，不是舍入决策点）。
from settlement_rules import money

# 口径绑定（复核链路 koujing_sha256）：口径哈希算法只有 koujing.digest 一份实现。
from koujing import digest as _koujing_digest


SCHEMA = "dealer-ops/settlement-result@1"


def load_settlement_result(workbook_path, workbook_sha):
    """同目录结果存在时必须绑定当前工作簿；过期结果不得静默沿用。"""
    path = os.path.join(os.path.dirname(os.path.abspath(workbook_path)), "_结算结果.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as handle:
            document = json.load(handle)
    except (OSError, ValueError) as exc:
        raise ValueError(f"无法读取结算结果凭证：{exc}") from exc
    if not isinstance(document, dict) or document.get("schema") != SCHEMA:
        raise ValueError("_结算结果.json schema 无效")
    if not isinstance(document.get("rule_result"), dict):
        raise ValueError("_结算结果.json 缺少 rule_result")
    recorded = document.get("workbook") or {}
    if os.path.realpath(recorded.get("path") or "") != os.path.realpath(workbook_path):
        raise ValueError("结算结果绑定的是另一份工作簿，已拒绝沿用")
    if recorded.get("sha256") != workbook_sha:
        raise ValueError("工作簿在结算后已变更，_结算结果.json 已过期；请重新出账或同步附表哈希")
    if document.get("business_type") != document["rule_result"].get("business_type"):
        raise ValueError("结算结果顶层业务类型与 rule_result 不一致")
    if document.get("business_type") not in ("consignment", "monthly_buyout"):
        raise ValueError("结算结果业务类型无效")
    return document


def validate_settlement_inputs(document, master, sources):
    """复核时的主档／源表必须与出账快照完全一致，禁止旧结果拼接新哈希。"""
    expected_master = document.get("master")
    if not isinstance(expected_master, dict) or not expected_master.get("path") \
            or not expected_master.get("sha256"):
        raise ValueError("结算结果缺少出账时的主档路径／哈希，请重新出账")
    if not isinstance(master, dict):
        raise ValueError("复核必须提供出账时的主档，不能用 null 跳过主档一致性校验")
    if os.path.realpath(master.get("path") or "") != os.path.realpath(expected_master["path"]):
        raise ValueError("本次复核主档与出账时使用的主档不是同一文件")
    if master.get("sha256") != expected_master["sha256"]:
        raise ValueError("主档在出账后已变更，原结算结果已过期；请重新出账并复核")

    # 口径绑定：铺货链路出账必带口径凭证（load_master 强制问卡），复核时口径文件
    # 必须与出账时一致——堵住「出账后、复核前改口径」窗口；月结不加载口径，缺席即跳过。
    if document["business_type"] == "consignment":
        expected_kj = document.get("koujing")
        if not isinstance(expected_kj, dict) or not expected_kj.get("path") \
                or not expected_kj.get("sha256"):
            raise ValueError("结算结果缺少口径凭证（koujing），请重新出账")
        if not os.path.isfile(expected_kj["path"]):
            raise ValueError("出账时的口径确认文件已不存在（%s）；请重新出账并复核"
                             % expected_kj["path"])
        try:
            with open(expected_kj["path"], encoding="utf-8") as fh:
                record = json.load(fh)
        except (OSError, ValueError) as exc:
            raise ValueError("口径确认文件无法读取（%s）：%s；请重新出账并复核"
                             % (expected_kj["path"], exc))
        try:
            current_kj_sha = _koujing_digest(record)
        except Exception as exc:  # noqa: BLE001 - 记录结构异常时给出可操作的中文报错
            raise ValueError("口径确认文件内容异常（%s）；请重走问卡门禁并重新出账" % exc)
        if current_kj_sha != expected_kj["sha256"]:
            raise ValueError("口径在出账后已变更（_口径确认.json 与出账凭证哈希不一致）；"
                             "改口径必须重走问卡门禁并重新出账，不得沿用原结算结果")

    recorded = document.get("sources")
    if not isinstance(recorded, dict):
        raise ValueError("结算结果缺少出账时的源文件路径／哈希，请重新出账")
    roles = ("receipts", "returns", "sales") if document["business_type"] == "monthly_buyout" \
        else ("jia", "sales", "inventory")
    current = {item.get("role"): item for item in sources if isinstance(item, dict)}
    for role in roles:
        expected = recorded.get(role)
        actual = current.get(role)
        if not isinstance(expected, dict) or not expected.get("path") or not expected.get("sha256"):
            raise ValueError(f"结算结果缺少源文件凭证（role={role}），请重新出账")
        if not isinstance(actual, dict) or not actual.get("path") or not actual.get("sha256"):
            raise ValueError(f"复核包无法读取源文件（role={role}），不能沿用原结算结果")
        if os.path.realpath(actual["path"]) != os.path.realpath(expected["path"]):
            raise ValueError(f"源文件路径在出账后已变更（role={role}），请重新出账")
        if actual["sha256"] != expected["sha256"]:
            raise ValueError(f"源文件在出账后已变更（role={role}），原结算结果已过期；请重新出账并复核")


def _number(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        number = Decimal(str(value))
    except Exception:
        return None
    if not number.is_finite():
        return None
    return int(number) if number == int(number) else float(number)


def _total(rows, key):
    values = [row.get(key) for row in rows if row.get(key) is not None]
    if not values:
        return None
    total = sum((Decimal(str(value)) for value in values), Decimal("0"))
    return _number(money(total))


def _name_map(workbook):
    names = {}
    for ws in workbook.worksheets:
        header_row, bar_col, name_col = None, None, None
        for row_index in range(1, min(ws.max_row, 200) + 1):
            headers = [str(ws.cell(row_index, col).value or "").replace("\n", "").strip()
                       for col in range(1, ws.max_column + 1)]
            bar_col = next((index + 1 for index, text in enumerate(headers)
                            if "商品条码" in text or text == "条码"), None)
            name_col = next((index + 1 for index, text in enumerate(headers)
                             if "商品名称" in text or "标准品名" in text), None)
            if bar_col and name_col:
                header_row = row_index
                break
        if not header_row:
            continue
        for row_index in range(header_row + 1, ws.max_row + 1):
            raw_bar = ws.cell(row_index, bar_col).value
            raw_name = ws.cell(row_index, name_col).value
            if isinstance(raw_bar, float) and raw_bar.is_integer():
                raw_bar = int(raw_bar)
            barcode = str(raw_bar or "").strip()
            name = str(raw_name or "").strip()
            if re.fullmatch(r"[0-9]{8,14}", barcode) and name and name != "合计":
                names.setdefault(barcode, name)
    return names


def _priced(rows, role, names):
    output = []
    for source in rows:
        row = dict(source)
        row["role"] = role
        row["name"] = names.get(str(row.get("barcode") or ""))
        output.append(row)
    return output


def _decal(rows, role, names):
    output = []
    for source in rows:
        output.append({
            "role": role, "id": source.get("id"), "date": source.get("date"),
            "bar": source.get("barcode"), "name": names.get(str(source.get("barcode") or "")),
            "settled_qty": _number(source.get("quantity")),
            "standard": _number(source.get("decal_standard")),
            "amount": _number(source.get("decal_amount")),
            "price_evidence": source.get("price_evidence"),
        })
    return output


def project_settlement(document, workbook):
    """返回 figures / sku_rows / priced_rows / tiehua_rows / warnings。"""
    business_type = document["business_type"]
    rule = document["rule_result"]
    names = _name_map(workbook)
    warnings = []
    if business_type == "consignment":
        lines = rule.get("priced_lines") or []
        quantities = rule.get("quantity_reconciliation") or []
        sku_rows = []
        for source in quantities:
            barcode = str(source.get("barcode") or "")
            own_lines = [line for line in lines if str(line.get("barcode") or "") == barcode]
            sku_rows.append({
                "bar": barcode, "name": names.get(barcode),
                "opening_qty": _number(source.get("opening_qty")),
                "settled_qty": _number(source.get("net_sales_qty")),
                "settled_amount": _total(own_lines, "amount"),
                "diff_qty": _number(source.get("variance_qty")), "unit_price": None,
                "theoretical_end_qty": _number(source.get("theoretical_end_qty")),
                "inventory_qty": _number(source.get("inventory_qty")),
                "addback_qty": _number(source.get("addback_qty")),
                "adjusted_actual_qty": _number(source.get("adjusted_actual_qty")),
                "variance_display_amount": _number(source.get("variance_display_amount")),
                "status": source.get("status"),
            })
        if any(row.get("variance_qty") is None for row in quantities):
            warnings.append("部分 SKU 尚待乙方库存，figures.diff_qty/diff_amount 只汇总已提供库存的行")
        figures = {
            "opening_qty": _total(quantities, "opening_qty"), "opening_amount": None,
            "settled_qty": _total(quantities, "net_sales_qty"),
            "settled_amount": _number(rule.get("settlement_amount")),
            "diff_qty": _total([row for row in quantities if row.get("variance_qty") is not None], "variance_qty"),
            "diff_amount": _number(rule.get("variance_display_amount")),
            "addback_qty": _total(quantities, "addback_qty"), "addback_amount": None,
            "tiehua_total": _number(rule.get("decal_amount")),
            "gross_sales_amount": _number(rule.get("gross_sales_amount")),
            "return_amount": _number(rule.get("return_amount")),
            "sales_settlement_amount": _number(rule.get("sales_settlement_amount")),
            "variance_display_amount": _number(rule.get("variance_display_amount")),
            "variance_included": rule.get("variance_included"),
            "period_end": rule.get("period_end"),
        }
        return {"figures": figures, "sku_rows": sku_rows,
                "priced_rows": _priced(lines, "铺货流水", names),
                "tiehua_rows": _decal(lines, "铺货销售／退货", names), "warnings": warnings}

    receipts = rule.get("priced_receipts") or []
    returns = rule.get("priced_returns") or []
    sales = rule.get("priced_sales_for_decal") or []
    barcodes = sorted({str(row.get("barcode") or "") for row in receipts + returns + sales})
    sku_rows = []
    for barcode in barcodes:
        own_receipts = [row for row in receipts if str(row.get("barcode") or "") == barcode]
        own_returns = [row for row in returns if str(row.get("barcode") or "") == barcode]
        own_sales = [row for row in sales if str(row.get("barcode") or "") == barcode]
        receipt_qty = _total(own_receipts, "quantity") or 0
        return_qty_signed = _total(own_returns, "quantity") or 0
        receipt_amount = _total(own_receipts, "amount") or 0
        return_amount = _total(own_returns, "amount") or 0
        sku_rows.append({
            "bar": barcode, "name": names.get(barcode), "opening_qty": None,
            "settled_qty": _number(Decimal(str(receipt_qty)) + Decimal(str(return_qty_signed))),
            "settled_amount": _number(Decimal(str(receipt_amount)) + Decimal(str(return_amount))),
            "diff_qty": None, "unit_price": None,
            "receipt_qty": receipt_qty, "return_qty": _number(-Decimal(str(return_qty_signed))),
            "sales_qty_for_decal_only": _total(own_sales, "quantity") or 0,
            "decal_amount": _number(Decimal(str(_total(own_sales, "decal_amount") or 0))
                                    + Decimal(str(_total(own_returns, "decal_amount") or 0))),
        })
    figures = {
        "opening_qty": None, "opening_amount": None,
        "settled_qty": _number(rule.get("buyout_qty")),
        "settled_amount": _number(rule.get("buyout_amount")),
        "diff_qty": None, "diff_amount": None, "addback_qty": None, "addback_amount": None,
        "tiehua_total": _number(rule.get("decal_amount")),
        "receipt_qty": _number(rule.get("receipt_qty")), "return_qty": _number(rule.get("return_qty")),
        "sales_qty_for_decal_only": _number(rule.get("sales_qty_for_decal_only")),
        "gross_receipt_amount": _number(rule.get("gross_receipt_amount")),
        "return_amount": _number(rule.get("return_amount")),
        "buyout_date": rule.get("buyout_date"), "payment_due_date": rule.get("payment_due_date"),
        "period_end": rule.get("period_end"),
    }
    return {"figures": figures, "sku_rows": sku_rows,
            "priced_rows": _priced(receipts, "签收", names) + _priced(returns, "退货", names),
            "tiehua_rows": _decal(sales, "乙方动销", names) + _decal(returns, "退货冲减", names),
            "warnings": warnings}
