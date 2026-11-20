#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从两层主档生成可核对待办；不创建定时任务，不替用户执行合同责任。

法定节假日与调休日历 JSON：
{"years":[2026], "holidays":["2026-10-01"], "workdays":["2026-10-10"]}
workdays 是调休上班日，优先于周末／节假日。没有覆盖相应年份的日历时，
工作日待办只给合同期限与估算标记，不伪造精确日期。付款截止日始终保留
合同原值，只有合同明确登记顺延／提前规则才计算调整；付款日文本无法
解析出明确日时只生成「资料缺失」待办与告警，不推定 30 天等默认日期。
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sys

import openpyxl

import workdays

from schema_common import sha256_file

SCHEMA = "dealer-ops/todos@1"


def fail(message):
    raise SystemExit("[中止] " + str(message))


def day(value, label="日期"):
    try:
        return workdays.parse_day(value, label)
    except ValueError as exc:
        fail(exc)


def optional_day(value, label="", warnings=None):
    if value is None or str(value).strip() == "":
        return None
    try:
        return workdays.parse_day(value, label)
    except ValueError:
        if warnings is not None and label:
            warnings.append(f"{label} 不是合法日期：{value!r}，按缺失处理")
        return None


def read_sheet(wb, name):
    if name not in wb.sheetnames:
        return []
    ws = wb[name]
    headers = [str(cell.value or "").strip() for cell in ws[1]]
    return [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)
            if any(value not in (None, "") for value in row)]


def load_calendar(path):
    try:
        return workdays.load_calendar(path)
    except workdays.CalendarError as exc:
        fail(exc)


def is_workday(value, calendar):
    return workdays.is_workday(value, calendar)


def workday_due(start, count, calendar):
    """返回 (日期或 None, 是否估算)。"""
    return workdays.workday_due(start, count, calendar)


def business_type_of(contract):
    """业务类型归一化：合同登记含「月结」即按月结口径处理。"""
    return str(contract.get("业务类型") or contract.get("结算方式") or "").strip()


def adjust_payment(due, rule, calendar):
    text = str(rule or "").strip()
    if not text or not re.search(r"顺延|提前|工作日|法定节假日|节假日", text):
        return due, False, "按合同登记日期，不自行顺延或提前"
    if re.search(r"(?:不|无需|不得)(?:需)?(?:顺延|提前)", text):
        return due, False, "合同明确不顺延／不提前，保留原截止日"
    if calendar is None:
        return None, True, "合同有工作日／节假日调整规则，但缺法定节假日及调休日历"
    direction = -1 if "提前" in text and "顺延" not in text else 1
    cursor = due
    while not is_workday(cursor, calendar):
        cursor += dt.timedelta(days=direction)
    if not all(year in calendar["years"]
               for year in range(min(due.year, cursor.year), max(due.year, cursor.year) + 1)):
        return None, True, "付款日按合同规则调整后超出日历覆盖年份，不生成伪精确日期"
    return cursor, False, f"按合同规则调整；合同原截止日 {due.isoformat()} 保留不变"


def digest_file(path):
    return sha256_file(path)


def generate(master, period, calendar=None):
    wb = openpyxl.load_workbook(master, data_only=True)
    try:
        contracts, terms = read_sheet(wb, "合同"), read_sheet(wb, "履约条款")
    finally:
        wb.close()
    start = day(period[:4] + "-" + period[4:] + "-01", "对账期间")
    next_month = dt.date(start.year + (start.month == 12), 1 if start.month == 12 else start.month + 1, 1)
    items, warnings = [], []

    def add(contract_no, kind, title, due=None, estimated=False, evidence="", source="", **extra):
        # 同一合同同类型可能同时存在季度、年度等不同关键值；来源与原文也参与去重。
        raw_id = "|".join((str(contract_no), period, kind, title, str(due),
                           str(source or ""), str(extra.get("key_value") or ""), str(evidence or "")))
        items.append({"id": "TODO-" + hashlib.sha256(raw_id.encode("utf-8")).hexdigest()[:12],
                      "contract_no": str(contract_no or ""), "period": period, "type": kind,
                      "title": title, "due_date": due.isoformat() if due else None,
                      "estimated": bool(estimated), "evidence": str(evidence or "")[:400],
                      "source": str(source or "")[:200], "status": "待处理", **extra})

    for index, contract in enumerate(contracts, 2):
        number = str(contract.get("合同编号") or "").strip()
        if not number:
            warnings.append(f"合同表第 {index} 行缺合同编号，相关待办未生成")
            continue
        contract_start = optional_day(contract.get("起始"), f"{number} 的起始日", warnings)
        if contract_start is not None and contract_start >= next_month:
            continue
        end = optional_day(contract.get("终止"), f"{number} 的终止日", warnings)
        if end is not None and end < start:
            continue
        source = f"{contract.get('来源文件') or '主档·合同'} {contract.get('来源位置') or ('第 ' + str(index) + ' 行')}"
        buyout = optional_day(contract.get("买断日"), f"{number} 的买断日", warnings)
        payment = optional_day(contract.get("付款截止日"), f"{number} 的付款截止日", warnings)
        if buyout:
            add(number, "月结买断", "核对本期签收与退货并准备月结买断对账", buyout, evidence=contract.get("原文依据"), source=source)
        pay_text = str(contract.get("付款日") or "").strip()
        if payment is None and pay_text:
            match = re.search(r"([0-9]{1,2})\s*日", pay_text)
            if match and 1 <= int(match.group(1)) <= 28:
                payment = next_month.replace(day=int(match.group(1)))
            else:
                warnings.append(f"{number} 的付款日「{pay_text}」无法解析出明确日"
                                f"（须为 1–28 日的月结日），不推定默认天数")
        if payment:
            adjusted, estimated, note = adjust_payment(payment, contract.get("付款顺延规则"), calendar)
            add(number, "付款截止", "核对对账确认与付款截止事项", adjusted, estimated,
                evidence=contract.get("原文依据") or pay_text, source=source,
                contract_due_date=payment.isoformat(), calculation_note=note)
        elif pay_text or "月结" in business_type_of(contract):
            add(number, "资料缺失", "补登记明确付款截止日（付款日无法确定，不默认日期）",
                None, True, evidence=pay_text, source=source)
        if end:
            add(number, "合同到期", "复核合同到期、续签或终止安排", end,
                evidence=contract.get("原文依据"), source=source)
        for field, kind in (("异议期(工作日)", "对账异议期"), ("签回期(工作日)", "对账单签回期")):
            raw = contract.get(field)
            if raw in (None, ""):
                continue
            try:
                count = int(raw)
                if count < 0 or str(raw).strip() != str(count):
                    raise ValueError
            except (TypeError, ValueError):
                warnings.append(f"{number} 的 {field} 不是非负整数，未计算工作日")
                continue
            due, estimated = workday_due(next_month, count, calendar)
            add(number, kind, f"按合同 {count} 个工作日期限复核{kind}", due, estimated,
                evidence=contract.get("原文依据"), source=source,
                basis_date=next_month.isoformat(), working_days=count,
                calculation_note="基准日暂取对账期末次日；实际起算以合同约定的发送／签收事实为准，无日历不生成精确日期")

    for index, term in enumerate(terms, 2):
        if str(term.get("提醒状态") or "").strip() in ("已完成", "已关闭", "无需提醒"):
            continue
        number = str(term.get("合同编号") or "").strip()
        kind = str(term.get("条款类型") or "履约条款").strip()
        value = str(term.get("关键值") or "").strip()
        evidence = str(term.get("原文依据") or "").strip()
        if not number or not value or not evidence:
            warnings.append(f"履约条款第 {index} 行缺合同编号／关键值／原文依据，未生成待办")
            continue
        add(number, "履约提醒", f"复核{kind}（不自动算返利、扣款或认定违约）", None, True,
            evidence=evidence, source=f"{term.get('来源文件') or '主档·履约条款'} {term.get('来源位置') or ('第 ' + str(index) + ' 行')}",
            key_value=value[:200])
    items.sort(key=lambda item: (item["due_date"] is None, item["due_date"] or "9999-12-31", item["id"]))
    return items, warnings


def main():
    parser = argparse.ArgumentParser(description="从合同主档与履约提醒层生成待办")
    parser.add_argument("master")
    parser.add_argument("--period", required=True, help="YYYYMM")
    parser.add_argument("--calendar", help="法定节假日与调休日历 JSON")
    parser.add_argument("--out", help="默认 <主档目录>/_待办.json")
    args = parser.parse_args()
    if not re.fullmatch(r"\d{6}", args.period) or not 1 <= int(args.period[4:]) <= 12:
        fail("--period 须为合法 YYYYMM")
    calendar = load_calendar(args.calendar)
    items, warnings = generate(args.master, args.period, calendar)
    output = os.path.abspath(args.out or os.path.join(os.path.dirname(os.path.abspath(args.master)), "_待办.json"))
    result = {"schema": SCHEMA, "master": {"path": os.path.abspath(args.master), "sha256": digest_file(args.master)},
              "period": args.period, "calendar": {"provided": bool(args.calendar),
              "covered_years": sorted(calendar["years"]) if calendar else []},
              "items": items, "warnings": warnings,
              "summary": {"total": len(items), "estimated": sum(item["estimated"] for item in items)},
              "note": "仅生成待办，不自动执行返利、保证金、违约／窜货责任处理，也未创建定时任务。"}
    os.makedirs(os.path.dirname(output), exist_ok=True)
    temp = output + ".tmp"
    with open(temp, "w", encoding="utf-8") as handle:
        json.dump(result, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    os.replace(temp, output)
    print(json.dumps({"ok": True, "out": output, "summary": result["summary"], "warnings": warnings}, ensure_ascii=False))


if __name__ == "__main__":
    main()
