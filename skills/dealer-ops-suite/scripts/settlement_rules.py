#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""经销结算的确定性规则核心。

只接受已经确认的合同价格、贴花标准和业务流水，不猜测缺失价格。
金额全部使用 Decimal；业务日期决定取价版本；退货必须能追溯原销售价。
本模块不读写 Excel，便于单元测试和两条对账链路共同复用。
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Iterable, Mapping, Optional

# 条码归一化的语义权威已收敛到 schema_common.normalize_barcode（共享内核）；
# 此处只再导出并保留公开名 barcode_of，两条对账引擎与规则核心继续用旧名调用。
from schema_common import normalize_barcode as barcode_of


CENT = Decimal("0.01")
ZERO = Decimal("0")
BUSINESS_TYPES = ("consignment", "monthly_buyout", "direct_buyout")


class RuleViolation(ValueError):
    """可机器处理的业务阻断项。"""

    def __init__(self, code: str, message: str, **context: Any):
        super().__init__(message)
        self.code = code
        self.message = message
        self.context = context

    def to_dict(self) -> dict[str, Any]:
        return {"code": self.code, "message": self.message, "context": self.context}


def dec(value: Any, label: str = "数值") -> Decimal:
    if value is None or isinstance(value, bool) or (isinstance(value, str) and not value.strip()):
        raise RuleViolation("INVALID_NUMBER", f"{label}缺失或不是有效数字：{value!r}", value=value)
    try:
        result = value if isinstance(value, Decimal) else Decimal(str(value).strip())
    except (InvalidOperation, ValueError) as exc:
        raise RuleViolation("INVALID_NUMBER", f"{label}不是有效数字：{value!r}", value=value) from exc
    if not result.is_finite():
        raise RuleViolation("INVALID_NUMBER", f"{label}不得为 NaN 或无穷大", value=value)
    if result == result.to_integral_value():
        result = result.quantize(Decimal(1))  # 60.0 → 60，JSON 序列化不带多余小数点
    return result


def money(value: Any) -> Decimal:
    return dec(value).quantize(CENT, rounding=ROUND_HALF_UP)


def _q(value):
    """数量归一：整数值不带小数指数（Decimal('0.0') → Decimal('0')），实数原样保留。"""
    if isinstance(value, Decimal) and value == value.to_integral_value():
        return value.quantize(Decimal(1))
    return value


def positive_qty(value: Any, label: str = "数量") -> Decimal:
    result = dec(value, label)
    if result < 0:
        raise RuleViolation("NEGATIVE_QUANTITY", f"{label}不得为负数；退货须显式标记类型", value=str(result))
    return result


def parse_day(value: Any, label: str = "业务日期") -> date:
    if value is None or (isinstance(value, str) and not value.strip()):
        raise RuleViolation("MISSING_BUSINESS_DATE", f"{label}缺失，已停止金额计算", value=value)
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise RuleViolation("INVALID_BUSINESS_DATE", f"{label}无法识别：{value!r}", value=value)


def normalize_business_type(value: Any) -> str:
    text = str(value or "").strip().lower().replace("-", "_").replace(" ", "")
    aliases = {
        "consignment": "consignment", "铺货代销": "consignment", "实销实结": "consignment",
        "铺货代销·实销实结": "consignment", "monthly_buyout": "monthly_buyout",
        "月结": "monthly_buyout", "月结买断": "monthly_buyout", "direct_buyout": "direct_buyout",
        "买断": "direct_buyout", "直接买断": "direct_buyout",
    }
    if text not in aliases:
        raise RuleViolation("UNKNOWN_BUSINESS_TYPE", f"无法识别业务类型：{value!r}；必须显式选择铺货代销、月结买断或直接买断")
    result = aliases[text]
    if result == "direct_buyout":
        raise RuleViolation("DIRECT_BUYOUT_NOT_RECONCILED", "直接买断在发货签收时完成销售，不进入铺货／月结对账引擎")
    return result


def rate_value_column(headers: Iterable[Any], aliases: Iterable[str], sheet_name: str) -> int:
    """两种对账引擎共用的价格／贴花列定位。

    只接受去掉空白及括号单位后的完整表头匹配，避免命中“原结算价”等
    其他业务列。多个价格列同时存在时无法判断权威来源，必须阻断。
    """
    import re

    def normalized(value: Any) -> str:
        text = re.sub(r"\s+", "", str(value or ""))
        return re.sub(r"[(（][^)）]*[)）]", "", text)

    candidates = [(index, normalized(header)) for index, header in enumerate(headers)
                  if normalized(header)]
    hits = set()
    for alias in aliases:
        target = normalized(alias)
        hits.update(index for index, header in candidates if header == target)
    if len(hits) > 1:
        actual = "、".join(str(list(headers)[index]) for index in sorted(hits))
        raise RuleViolation(
            "AMBIGUOUS_RATE_COLUMN",
            f"主档「{sheet_name}」同时存在多个价格／标准列：{actual}；不得任选一列，请只保留一个权威列",
            columns=sorted(hits),
        )
    if not hits:
        raise RuleViolation(
            "MISSING_RATE_COLUMN",
            f"主档「{sheet_name}」缺少列：{'／'.join(aliases)}",
        )
    return next(iter(hits))


@dataclass(frozen=True)
class Rate:
    barcode: str
    value: Decimal
    start: date
    end: Optional[date]
    source: str = ""

    def covers(self, day: date) -> bool:
        return self.start <= day and (self.end is None or day <= self.end)


class RateBook:
    """同一 SKU 的 SCD2 价格／贴花区间，重叠立即阻断。"""

    def __init__(self, rows: Iterable[Mapping[str, Any]], kind: str = "合同价"):
        self.kind = kind
        self._rates: dict[str, list[Rate]] = {}
        for index, row in enumerate(rows, 2):
            barcode = barcode_of(row.get("barcode", row.get("商品条码")))
            if not barcode:
                raise RuleViolation("MISSING_BARCODE", f"{kind}第 {index} 行缺少商品条码", row=index)
            raw_value = row.get("value", row.get("rate", row.get("供货价", row.get("贴花标准"))))
            value = dec(raw_value, f"{kind}·条码 {barcode}")
            if value < 0:
                raise RuleViolation("NEGATIVE_RATE", f"{kind}不得为负数：条码 {barcode}", barcode=barcode)
            start = parse_day(row.get("start", row.get("生效起")), f"{kind}生效起")
            raw_end = row.get("end", row.get("生效止"))
            end = parse_day(raw_end, f"{kind}生效止") if raw_end not in (None, "") else None
            if end is not None and end < start:
                raise RuleViolation("INVALID_RATE_INTERVAL", f"{kind}生效止早于生效起：条码 {barcode}", barcode=barcode)
            rate = Rate(barcode, value, start, end, str(row.get("source", row.get("依据文件", "")) or ""))
            self._rates.setdefault(barcode, []).append(rate)
        for barcode, rates in self._rates.items():
            rates.sort(key=lambda item: (item.start, item.end or date.max))
            for previous, current in zip(rates, rates[1:]):
                if previous.end is None or current.start <= previous.end:
                    raise RuleViolation(
                        "OVERLAPPING_RATE_INTERVALS",
                        f"条码 {barcode} 的{kind}生效区间重叠：{previous.start}—{previous.end or '长期'} 与 {current.start}—{current.end or '长期'}",
                        barcode=barcode,
                    )

    def rate_on(self, barcode: Any, day: Any) -> Rate:
        bc = barcode_of(barcode)
        business_day = parse_day(day)
        hits = [rate for rate in self._rates.get(bc, []) if rate.covers(business_day)]
        if len(hits) != 1:
            raise RuleViolation(
                "NO_UNIQUE_RATE",
                f"条码 {bc} 在 {business_day} 没有唯一生效的{self.kind}（命中 {len(hits)} 条）；已停止金额计算",
                barcode=bc,
                date=business_day.isoformat(),
                hits=len(hits),
            )
        return hits[0]

    def validate_period(self, barcode: Any, start: Any, end: Any) -> None:
        """要求指定 SKU 在整个期间连续有效，断档不得靠期末价补猜。"""
        bc = barcode_of(barcode)
        first, last = parse_day(start), parse_day(end)
        cursor = first
        rates = self._rates.get(bc, [])
        while cursor <= last:
            hits = [rate for rate in rates if rate.covers(cursor)]
            if len(hits) != 1:
                raise RuleViolation(
                    "RATE_COVERAGE_GAP",
                    f"条码 {bc} 的{self.kind}在 {cursor} 出现断档或重叠；已停止金额计算",
                    barcode=bc,
                    date=cursor.isoformat(),
                )
            hit = hits[0]
            cursor = (hit.end or last) + timedelta(days=1)


def _line_id(line: Mapping[str, Any], index: int) -> str:
    return str(line.get("id", line.get("serial", line.get("流水号", index))) or index).strip()


def _line_kind(line: Mapping[str, Any], default: str) -> str:
    kind = str(line.get("kind", line.get("type", line.get("业务类型", default))) or default).strip().lower()
    aliases = {"销售": "sale", "动销": "sale", "退货": "return", "冲销": "return", "签收": "receipt"}
    return aliases.get(kind, kind)


def _original_ref(line: Mapping[str, Any], preferred: str) -> str:
    return str(line.get(preferred) or line.get("original_sale_id") or line.get("原销售流水号") or "").strip()


def _original_price(
    line: Mapping[str, Any],
    original: Optional[Mapping[str, Any]],
    prices: RateBook,
) -> tuple[Decimal, str]:
    if original is not None:
        line_barcode = barcode_of(line.get("barcode", line.get("商品条码")))
        original_barcode = barcode_of(original.get("barcode", original.get("商品条码")))
        if line_barcode != original_barcode:
            raise RuleViolation("ORIGINAL_BARCODE_MISMATCH", "退货条码与原流水条码不一致，已停止金额计算",
                                return_barcode=line_barcode, original_barcode=original_barcode)
    explicit = line.get("original_unit_price", line.get("原销售单价", line.get("原签收单价")))
    if explicit not in (None, ""):
        return dec(explicit, "原交易单价"), "退货行明确登记的原交易单价"
    if original is None:
        raise RuleViolation(
            "ORIGINAL_PRICE_REQUIRED",
            "退货无法追溯原销售流水或明确原销售价，已停止金额计算；不得按退货日价格猜算",
            line=dict(line),
        )
    original_explicit = original.get("unit_price", original.get("原销售单价"))
    if original_explicit not in (None, ""):
        return dec(original_explicit, "原销售单价"), "原销售流水登记的单价"
    original_day = parse_day(original.get("date", original.get("业务日期")), "原销售日期")
    original_barcode = original.get("barcode", original.get("商品条码"))
    return prices.rate_on(original_barcode, original_day).value, "按原销售日期匹配合同价"


def _first_present(*values: Any) -> Any:
    """取第一个非空（非 None、非空白字符串）值；引擎常显式传 None，不能挡住后续兜底。"""
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return None


def _original_decal(
    line: Mapping[str, Any],
    original: Optional[Mapping[str, Any]],
    decals: RateBook,
    current_period_standard: Optional[Decimal] = None,
) -> tuple[Decimal, bool, str]:
    """退货贴花冲减依据的优先级：

    1. 退货行明确登记的标准／计提状态；
    2. 可追溯原销售行的实际计提标准（本期原销售按实际计提标准；
       跨期历史流水按登记字段，缺失才按原销售日期政策兜底）；
    3. 证据不足一律阻断，不默认冲减或不冲减。
    """
    source = original if original is not None else line
    charged = _first_present(line.get("original_decal_charged"),
                             source.get("decal_charged"), source.get("原销售已计提贴花"))
    if original is not None and barcode_of(line.get("barcode", line.get("商品条码"))) != barcode_of(
            original.get("barcode", original.get("商品条码"))):
        raise RuleViolation("ORIGINAL_BARCODE_MISMATCH", "退货条码与原销售贴花流水条码不一致，已停止贴花冲减")
    if isinstance(charged, str):
        normalized = charged.strip().lower()
        if normalized in ("是", "已计提", "true", "1", "有"):
            charged = True
        elif normalized in ("否", "未计提", "false", "0", "无"):
            charged = False
    if charged is None and current_period_standard is not None:
        # 本期原销售确定已按期末标准计提，退货按同一标准冲减，与货款口径相互独立。
        charged = True
    if charged not in (True, False, 0, 1) or isinstance(charged, str):
        raise RuleViolation(
            "ORIGINAL_DECAL_UNKNOWN",
            "退货是否已计提贴花无法确认，已列为待补依据；不得默认冲减或不冲减",
            line=dict(line),
        )
    if not bool(charged):
        return ZERO, False, "原销售未计提贴花"
    standard = _first_present(line.get("original_decal_standard"),
                              source.get("decal_standard"), source.get("原销售贴花标准"),
                              current_period_standard)
    if standard is not None:
        if standard is current_period_standard:
            return current_period_standard, True, "按本期原销售实际计提的贴花标准冲减"
        return dec(standard, "原销售贴花标准"), True, "原销售贴花标准有明确登记"
    if original is None:
        raise RuleViolation("ORIGINAL_DECAL_UNKNOWN", "退货缺少原销售贴花标准或原销售流水", line=dict(line))
    day = parse_day(original.get("date", original.get("业务日期")), "原销售日期")
    rate = decals.rate_on(original.get("barcode", original.get("商品条码")), day)
    return rate.value, True, "按原销售日期匹配贴花标准"


def _index_originals(rows: Iterable[Mapping[str, Any]], id_field: str) -> dict[str, Mapping[str, Any]]:
    result: dict[str, Mapping[str, Any]] = {}
    for index, row in enumerate(rows, 1):
        key = str(row.get(id_field) or row.get("id") or row.get("serial") or row.get("流水号") or "").strip()
        if not key:
            raise RuleViolation("MISSING_ORIGINAL_REFERENCE", f"原始流水第 {index} 行缺少可用于退货追溯的流水号")
        if key in result:
            raise RuleViolation("DUPLICATE_TRANSACTION_ID", f"原始流水号重复：{key}", transaction_id=key)
        result[key] = row
    return result


def _variance_authorized(period_start: date, period_end: date, authorization: Optional[Mapping[str, Any]]) -> bool:
    if not authorization:
        return False
    if authorization.get("approved") is not True:
        return False
    try:
        approved_start = parse_day(authorization.get("period_start"), "授权期间起")
        approved_end = parse_day(authorization.get("period_end"), "授权期间止")
    except RuleViolation as exc:
        raise RuleViolation("VARIANCE_AUTHORIZATION_REQUIRED", "盈亏纳入授权必须绑定明确的对账期间起止") from exc
    if approved_start != period_start or approved_end != period_end:
        raise RuleViolation(
            "VARIANCE_AUTHORIZATION_PERIOD_MISMATCH",
            "盈亏纳入授权仅对原授权期间有效，不得自动继承到本期",
            authorized_start=approved_start.isoformat(),
            authorized_end=approved_end.isoformat(),
            current_start=period_start.isoformat(),
            current_end=period_end.isoformat(),
        )
    return True


def calculate_consignment(
    *,
    period_start: Any,
    period_end: Any,
    opening: Mapping[str, Any],
    inventory: Mapping[str, Any],
    addback: Optional[Mapping[str, Any]],
    sales: Iterable[Mapping[str, Any]],
    prices: RateBook,
    decals: Optional[RateBook] = None,
    original_sales: Iterable[Mapping[str, Any]] = (),
    variance_authorization: Optional[Mapping[str, Any]] = None,
) -> dict[str, Any]:
    """铺货代销·实销实结：数量勾稽、金额分列。

    正盈亏数量为短缺，负数为盘盈。默认只列示盈亏，不进入结算总额。
    正常销售贴花沿用对账期末日标准；退货贴花按原销售标准冲减。
    """
    start, end = parse_day(period_start, "对账期间起"), parse_day(period_end, "对账期间止")
    if end < start:
        raise RuleViolation("INVALID_PERIOD", "对账期末日早于期初日")
    include_variance = _variance_authorized(start, end, variance_authorization)
    current_sales = list(sales)
    _index_originals(current_sales, "id")  # 本期销售／退货流水号不得重复
    current_sale_keys = _index_originals([x for x in current_sales if _line_kind(x, "sale") == "sale"], "id")
    originals = _index_originals(list(original_sales) + [x for x in current_sales if _line_kind(x, "sale") == "sale"], "id")
    priced_lines: list[dict[str, Any]] = []
    barcode_qty: dict[str, Decimal] = {}
    gross_sales_amount = ZERO
    return_amount = ZERO
    sale_decal = ZERO
    return_decal = ZERO

    for index, line in enumerate(current_sales, 1):
        kind = _line_kind(line, "sale")
        if kind not in ("sale", "return"):
            raise RuleViolation("UNKNOWN_TRANSACTION_TYPE", f"铺货代销流水只接受销售或退货：{kind!r}")
        bc = barcode_of(line.get("barcode", line.get("商品条码")))
        if not bc:
            raise RuleViolation("MISSING_BARCODE", f"销售流水第 {index} 行缺少条码")
        business_day = parse_day(line.get("date", line.get("业务日期")), "销售／退货日期")
        if not start <= business_day <= end:
            raise RuleViolation("TRANSACTION_OUTSIDE_PERIOD", f"流水 {business_day} 不在本期 {start}—{end} 内")
        qty = positive_qty(line.get("quantity", line.get("qty", line.get("数量"))))
        signed_qty = -qty if kind == "return" else qty
        line_id = _line_id(line, index)
        line_decal_standard = None
        line_decal_amount = ZERO
        if kind == "sale":
            rate = prices.rate_on(bc, business_day)
            unit_price, evidence = rate.value, rate.source or "按销售日期匹配合同价"
            amount = money(qty * unit_price)
            gross_sales_amount += amount
            if decals is not None:
                line_decal_standard = decals.rate_on(bc, end).value
                line_decal_amount = money(qty * line_decal_standard)
                sale_decal += line_decal_amount
            original_id = ""
        else:
            ref = _original_ref(line, "original_sale_id")
            original = originals.get(ref) if ref else None
            unit_price, evidence = _original_price(line, original, prices)
            amount = money(-qty * unit_price)
            return_amount += amount
            original_id = ref
            if decals is not None:
                # 本期原销售按期末标准实际计提；退货按同一标准冲减，期中调标不留残值。
                current_standard = (money(decals.rate_on(bc, end).value)
                                  if ref and ref in current_sale_keys else None)
                standard, charged, _ = _original_decal(line, original, decals, current_standard)
                if charged:
                    line_decal_standard = standard
                    line_decal_amount = money(-qty * standard)
                    return_decal += line_decal_amount
        barcode_qty[bc] = barcode_qty.get(bc, ZERO) + signed_qty
        priced_lines.append({
            "id": line_id, "barcode": bc, "kind": kind, "date": business_day.isoformat(),
            "quantity": signed_qty, "unit_price": money(unit_price), "amount": amount,
            "original_sale_id": original_id, "price_evidence": evidence,
            "decal_standard": money(line_decal_standard) if line_decal_standard is not None else None,
            "decal_amount": line_decal_amount,
        })

    all_barcodes = set(opening) | set(inventory) | set(addback or {}) | set(barcode_qty)
    sku_rows: list[dict[str, Any]] = []
    total_variance_amount = ZERO
    for bc in sorted(all_barcodes):
        opening_qty = dec(opening.get(bc, ZERO), f"条码 {bc} 甲方期初数量")
        net_qty = barcode_qty.get(bc, ZERO)
        theoretical_end = opening_qty - net_qty
        if bc not in inventory or inventory.get(bc) is None:
            actual = None
            adjusted = None
            shortage = None
            variance_amount = None
            status = "待补乙方库存"
        else:
            actual = dec(inventory[bc], f"条码 {bc} 库存导出数量")
            back = dec((addback or {}).get(bc, ZERO), f"条码 {bc} 导出日前已销回加")
            adjusted = actual + back
            shortage = theoretical_end - adjusted
            closing_rate = prices.rate_on(bc, end)
            variance_amount = money(shortage * closing_rate.value)
            total_variance_amount += variance_amount
            status = "账实相符" if shortage == 0 else ("短缺·待双方核对" if shortage > 0 else "盘盈·待双方核对")
        sku_rows.append({
            "barcode": bc, "opening_qty": _q(opening_qty), "net_sales_qty": _q(net_qty),
            "theoretical_end_qty": _q(theoretical_end), "inventory_qty": _q(actual),
            "addback_qty": _q(dec((addback or {}).get(bc, ZERO))), "adjusted_actual_qty": _q(adjusted),
            "variance_qty": _q(shortage), "variance_display_amount": variance_amount,
            "status": status,
        })

    sales_amount = gross_sales_amount + return_amount
    settlement_amount = sales_amount + (total_variance_amount if include_variance else ZERO)
    decal_amount = sale_decal + return_decal
    return {
        "business_type": "consignment", "period_start": start.isoformat(), "period_end": end.isoformat(),
        "quantity_reconciliation": sku_rows, "priced_lines": priced_lines,
        "gross_sales_amount": money(gross_sales_amount), "return_amount": money(return_amount),
        "sales_settlement_amount": money(sales_amount),
        "variance_display_amount": money(total_variance_amount),
        "variance_included": include_variance,
        "variance_authorization": "本期单次授权" if include_variance else "未纳入本期结算总额",
        "settlement_amount": money(settlement_amount),
        "decal_sale_amount": money(sale_decal), "decal_return_amount": money(return_decal),
        "decal_amount": money(decal_amount), "decal_pricing_rule": "正常销售取对账期末日标准；退货按原销售标准冲减",
        "amount_reconciliation": "数量勾稽、金额分列；不存在调价价差倒挤项",
        "opening_source": "本期甲方导出；不强制与上期台账期末衔接",
    }


def calculate_monthly_buyout(
    *,
    period_start: Any,
    period_end: Any,
    receipts: Iterable[Mapping[str, Any]],
    returns: Iterable[Mapping[str, Any]],
    sales: Iterable[Mapping[str, Any]],
    prices: RateBook,
    decals: Optional[RateBook] = None,
    original_sales: Iterable[Mapping[str, Any]] = (),
    original_receipts: Iterable[Mapping[str, Any]] = (),
    buyout_date: Any,
    payment_due_date: Any,
) -> dict[str, Any]:
    """月结买断：签收净额决定货款，乙方动销只决定贴花。"""
    start, end = parse_day(period_start, "对账期间起"), parse_day(period_end, "对账期间止")
    if end < start:
        raise RuleViolation("INVALID_PERIOD", "对账期末日早于期初日")
    buyout_day = parse_day(buyout_date, "买断日")
    payment_day = parse_day(payment_due_date, "付款截止日")
    receipt_rows = list(receipts)
    originals_receipts = _index_originals(list(original_receipts) + receipt_rows, "id")
    return_rows = list(returns)
    _index_originals(return_rows, "id")  # 本期退货流水号不得重复
    current_sales = list(sales)
    _index_originals(current_sales, "id")  # 本期销售流水号不得重复
    originals_sales = _index_originals(list(original_sales) + current_sales, "id")

    receipt_qty = ZERO
    gross_receipt_amount = ZERO
    priced_receipts: list[dict[str, Any]] = []
    for index, line in enumerate(receipt_rows, 1):
        bc = barcode_of(line.get("barcode", line.get("商品条码")))
        day = parse_day(line.get("date", line.get("签收日期")), "签收日期")
        if not start <= day <= end:
            raise RuleViolation("TRANSACTION_OUTSIDE_PERIOD", f"签收流水 {day} 不在本期内")
        qty = positive_qty(line.get("quantity", line.get("qty", line.get("签收数量"))))
        rate = prices.rate_on(bc, day)
        amount = money(qty * rate.value)
        receipt_qty += qty
        gross_receipt_amount += amount
        if not bc:
            raise RuleViolation("MISSING_BARCODE", f"签收流水第 {index} 行缺少条码")
        priced_receipts.append({
            "id": _line_id(line, index), "barcode": bc, "date": day.isoformat(),
            "quantity": qty, "unit_price": money(rate.value), "amount": amount,
            "price_evidence": rate.source or "按签收日期匹配合同价",
        })

    return_qty = ZERO
    return_amount = ZERO
    return_decal = ZERO
    priced_returns: list[dict[str, Any]] = []
    for index, line in enumerate(return_rows, 1):
        bc = barcode_of(line.get("barcode", line.get("商品条码")))
        day = parse_day(line.get("date", line.get("退货日期")), "退货日期")
        if not start <= day <= end:
            raise RuleViolation("TRANSACTION_OUTSIDE_PERIOD", f"退货流水 {day} 不在本期内")
        qty = positive_qty(line.get("quantity", line.get("qty", line.get("退货数量"))))
        receipt_ref = str(line.get("original_receipt_id") or line.get("原签收流水号") or "").strip()
        original_receipt = originals_receipts.get(receipt_ref) if receipt_ref else None
        unit_price, evidence = _original_price(line, original_receipt, prices)
        amount = money(-qty * unit_price)
        return_qty += qty
        return_amount += amount
        sale_ref = _original_ref(line, "original_sale_id")
        original_sale = originals_sales.get(sale_ref) if sale_ref else None
        if not bc:
            raise RuleViolation("MISSING_BARCODE", f"退货流水第 {index} 行缺少条码")
        decal_standard = None
        line_return_decal = ZERO
        if decals is not None:
            standard, charged, _ = _original_decal(line, original_sale, decals)
            if charged:
                decal_standard = standard
                line_return_decal = money(-qty * standard)
                return_decal += line_return_decal
        priced_returns.append({
            "id": _line_id(line, index), "barcode": bc, "date": day.isoformat(),
            "quantity": -qty, "unit_price": money(unit_price), "amount": amount,
            "original_receipt_id": receipt_ref, "original_sale_id": sale_ref, "price_evidence": evidence,
            "decal_standard": money(decal_standard) if decal_standard is not None else None,
            "decal_amount": line_return_decal,
        })

    sale_qty = ZERO
    sale_decal = ZERO
    priced_sales: list[dict[str, Any]] = []
    for index, line in enumerate(current_sales, 1):
        if _line_kind(line, "sale") != "sale":
            raise RuleViolation("UNKNOWN_TRANSACTION_TYPE", "月结乙方动销表只接受正向销售；退货请放退货明细")
        bc = barcode_of(line.get("barcode", line.get("商品条码")))
        if not bc:
            raise RuleViolation("MISSING_BARCODE", f"乙方动销第 {index} 行缺少条码")
        day = parse_day(line.get("date", line.get("动销日期")), "乙方动销日期")
        if not start <= day <= end:
            raise RuleViolation("TRANSACTION_OUTSIDE_PERIOD", f"乙方动销流水 {day} 不在本期内")
        qty = positive_qty(line.get("quantity", line.get("qty", line.get("动销数量"))))
        rate = decals.rate_on(bc, day) if decals is not None else None
        amount = money(qty * rate.value) if rate is not None else ZERO
        sale_qty += qty
        sale_decal += amount
        priced_sales.append({
            "id": _line_id(line, index), "barcode": bc, "date": day.isoformat(),
            "quantity": qty,
            "decal_standard": money(rate.value) if rate is not None else None,
            "decal_amount": amount,
            "price_evidence": (rate.source or "按乙方动销日期匹配贴花标准") if rate is not None else "未提供贴花政策，仅记录动销数量",
        })

    buyout_qty = receipt_qty - return_qty
    buyout_amount = gross_receipt_amount + return_amount
    return {
        "business_type": "monthly_buyout", "period_start": start.isoformat(), "period_end": end.isoformat(),
        "receipt_qty": receipt_qty, "return_qty": return_qty, "buyout_qty": buyout_qty,
        "sales_qty_for_decal_only": sale_qty,
        "gross_receipt_amount": money(gross_receipt_amount), "return_amount": money(return_amount),
        "buyout_amount": money(buyout_amount), "settlement_amount": money(buyout_amount),
        "decal_sale_amount": money(sale_decal), "decal_return_amount": money(return_decal),
        "decal_amount": money(sale_decal + return_decal),
        "priced_receipts": priced_receipts, "priced_returns": priced_returns, "priced_sales_for_decal": priced_sales,
        "buyout_date": buyout_day.isoformat(), "payment_due_date": payment_day.isoformat(),
        "quantity_rule": "买断数量＝本期签收数量－本期退货数量；乙方动销不减少买断货款",
        "price_rule": "买断货款按每笔签收日期取合同价；退货按原签收价冲减",
        "decal_pricing_rule": "贴花按每笔乙方动销日期取标准；退货按原销售标准冲减",
    }


def jsonable(value: Any) -> Any:
    """将规则结果递归转换为可写 JSON 的数值和字符串。"""
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Mapping):
        return {str(key): jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [jsonable(item) for item in value]
    return value
