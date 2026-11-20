#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""工作日计算单一权威实现（make_todos 与 ledger_read 共用）。

法定节假日与调休日历 JSON：
{"years":[2026], "holidays":["2026-10-01"], "workdays":["2026-10-10"]}
workdays 是调休上班日，优先于周末／节假日。没有覆盖相应年份的日历时，
工作日计算只给 None 与估算标记，不伪造精确日期。

口径约定：
- 提供日历 → 按法定节假日及调休日历计算；
- 未提供日历 → 仅排除周末（等价 Excel WORKDAY 无节假日参数），
  调用方必须在输出中标注该口径，不得与含节假日的结果混用同一语义。
"""
import datetime as dt
import json


class CalendarError(ValueError):
    """法定节假日及调休日历无法读取或内容非法。"""


def parse_day(value, label="日期"):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"{label}须为 YYYY-MM-DD 或 YYYY/MM/DD：{value!r}")


def load_calendar(path):
    """读取法定节假日与调休日历；path 为空返回 None。"""
    if not path:
        return None
    try:
        with open(path, encoding="utf-8") as handle:
            raw = json.load(handle)
        holidays = {parse_day(value, "节假日") for value in raw.get("holidays", [])}
        workdays = {parse_day(value, "调休上班日") for value in raw.get("workdays", [])}
        years = {int(value) for value in raw.get("years", [])}
    except (OSError, ValueError, TypeError, AttributeError) as exc:
        raise CalendarError("法定节假日／调休日历无法读取：" + str(exc)) from exc
    if not years:
        years = {value.year for value in holidays | workdays}
    if not years:
        raise CalendarError("空日历必须显式填写 years，否则无法判断是否覆盖相关年份")
    return {"holidays": holidays, "workdays": workdays, "years": years}


def is_workday(value, calendar):
    if value in calendar["workdays"]:
        return True
    if value in calendar["holidays"] or value.weekday() >= 5:
        return False
    return True


def workday_due(start, count, calendar):
    """含日历的工作日递推。返回 (日期或 None, 是否估算)。

    calendar 为 None 或未覆盖递推涉及的全部年份时返回 (None, True)，
    调用方据此标注估算，不得给出伪精确日期。
    """
    if calendar is None or not all(
            year in calendar["years"]
            for year in (start.year, (start + dt.timedelta(days=int(count) * 3 + 20)).year)):
        return None, True
    cursor, remaining = start, int(count)
    while remaining:
        cursor += dt.timedelta(days=1)
        if is_workday(cursor, calendar):
            remaining -= 1
    return cursor, False


def add_workdays(start, days, calendar=None):
    """统一工作日递推入口。返回 (日期或 None, 是否估算, 口径说明)。"""
    if calendar is not None:
        due, estimated = workday_due(start, days, calendar)
        if estimated:
            return None, True, "缺覆盖相关年份的法定节假日及调休日历，不生成精确日期"
        return due, False, "按法定节假日及调休日历计算工作日"
    cursor, remaining = start, int(days)
    while remaining:
        cursor += dt.timedelta(days=1)
        if cursor.weekday() < 5:
            remaining -= 1
    return cursor, True, "仅按周末计算工作日，不含法定节假日（未提供日历）"
