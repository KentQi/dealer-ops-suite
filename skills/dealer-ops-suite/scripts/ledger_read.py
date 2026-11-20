#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""台账「对账待办」只读查询。

回答：**这期该做啥、谁逾期、卡在哪。**

两个工程事实（已实测）：
  1. 台账的 J/K/L 等列是公式。文件若上次由脚本保存，openpyxl 的 data_only
     缓存为 None——直接读会得出"本月应做 0"的错误结论。
  2. 因此本脚本采用 **有缓存用缓存、无缓存按台账公式即时算** 的策略，
     并在输出里标明取值来源，绝不静默返回错值。

用法：
    python3 ledger_read.py --ledger "<台账.xlsx>"
    python3 ledger_read.py --ledger "<台账.xlsx>" --customer 示例经销商
    python3 ledger_read.py --ledger "<台账.xlsx>" --json

退出码：0 = 正常；1 = 无法解析
"""

import argparse
import json
import re
import sys
from datetime import date, datetime, timedelta

try:
    import openpyxl
except ImportError:
    print("需要 openpyxl：python3 -m pip install openpyxl", file=sys.stderr)
    sys.exit(1)

import workdays

SHEET_NAME = "对账待办"


# ---------- 通用 ----------

def norm(s):
    if s is None:
        return ""
    if isinstance(s, datetime):
        return s.strftime("%Y-%m-%d")
    if isinstance(s, date):
        return s.strftime("%Y-%m-%d")
    return re.sub(r"\s+", " ", str(s)).strip()


def to_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace("元", "").strip())
    except ValueError:
        return None


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    return None


# ---------- Excel 公式的 Python 等价实现 ----------

def add_workdays(start: date, days: int, calendar=None) -> date:
    """工作日递推，与 make_todos 共用同一权威实现。

    calendar=None 时等价 Excel WORKDAY(start, days)（仅排除周六日，与台账
    公式一致）；提供法定节假日及调休日历时按同一日历口径计算。
    日历未覆盖相关年份时返回 None，调用方不得给出伪精确日期。
    """
    due, _estimated, _note = workdays.add_workdays(start, days, calendar)
    return due


WORKDAY_BASIS_WITH_CALENDAR = "按法定节假日及调休日历计算工作日"
WORKDAY_BASIS_WEEKEND_ONLY = "仅按周末计算工作日，不含法定节假日（与台账 WORKDAY 公式一致）"


def period_to_next_month_first(period: str):
    """F 交单截止 = DATE(year, month+1, 1)"""
    m = re.fullmatch(r"(\d{4})-(\d{2})", period or "")
    if not m:
        return None
    y, mo = int(m.group(1)), int(m.group(2))
    mo += 1
    if mo == 13:
        y, mo = y + 1, 1
    return date(y, mo, 1)


def next_month_day(period: str, day: int):
    """I 回款截止 = DATE(year, month+1, 结款日)"""
    first = period_to_next_month_first(period)
    if not first or not day:
        return None
    try:
        return first.replace(day=int(day))
    except ValueError:
        return None


def compute_todo(q, r, i_date, p, o, n, m, f_date, today):
    """J 当前应做——逐条对应台账公式。"""
    if q is not None and q > 0 and r is not None and r >= q:
        return "已完成"
    if q is not None and q > 0 and i_date and today > i_date and (r is None or r < q):
        return "催回款（已逾期）"
    if p:
        return "催回款"
    if o:
        return "催乙方寄回原件"
    if n:
        return "出具对账单"
    if m:
        return "核算价格"
    if f_date and today > f_date:
        return "催乙方交单（已逾期）"
    return "等待乙方交单"


def compute_days_left(todo, i_date, h_date, g_date, f_date, today):
    """K 距截止"""
    if todo == "已完成":
        return None
    if todo in ("催回款", "催回款（已逾期）"):
        target = i_date
    elif todo == "催乙方寄回原件":
        target = h_date
    elif todo in ("出具对账单", "核算价格"):
        target = g_date
    else:
        target = f_date
    if not target:
        return None
    return (target - today).days


def compute_urgency(todo, days_left, q, i_date, r, today):
    """L 紧急度"""
    if todo == "已完成":
        return "已完成"
    if q is not None and q > 0 and i_date and today > i_date and (r is None or r < q):
        return "逾期"
    if days_left is None:
        return ""
    if days_left < 0:
        return "逾期"
    if days_left == 0:
        return "今日"
    if days_left <= 3:
        return "3日内"
    if days_left <= 7:
        return "7日内"
    return "正常"


# ---------- 表结构 ----------

def find_header_row(ws):
    for r in range(1, min(ws.max_row, 8) + 1):
        vals = [norm(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        if any("账期" in v for v in vals) and any("客户" in v for v in vals):
            return r
    return None


COL_RULES = {
    "id": ("待办编号",),
    "period": ("账期",),
    "customer": ("客户名称", "客户"),
    "contract": ("合同编号",),
    "pay_day": ("结款日",),
    "submit_due": ("交单截止",),
    "issue_due": ("出单截止",),
    "confirm_due": ("确认截止",),
    "collect_due": ("回款截止",),
    "todo": ("当前应做",),
    "days_left": ("距截止",),
    "urgency": ("紧急度",),
    "received": ("资料收到日",),
    "priced": ("核算完成日",),
    "issued": ("对账单出具日",),
    "original": ("原件到达日",),
    "amount": ("本期结算金额",),
    "collected": ("累计回款",),
    "blocker": ("卡点责任方",),
    "note": ("备注",),
}


def map_cols(ws, header_row):
    m = {}
    for c in range(1, ws.max_column + 1):
        h = norm(ws.cell(row=header_row, column=c).value)
        if not h:
            continue
        for key, aliases in COL_RULES.items():
            if key in m:
                continue
            if any(a in h for a in aliases):
                m[key] = c
    return m


def collect(ws_v, ws_f, header_row, cols, today, customer_filter=None, calendar=None):
    items = []
    for r in range(header_row + 1, ws_v.max_row + 1):
        def cached(k):
            return ws_v.cell(row=r, column=cols[k]).value if k in cols else None

        def formula(k):
            return ws_f.cell(row=r, column=cols[k]).value if k in cols else None

        period = norm(cached("period"))
        customer = norm(cached("customer"))
        # 待办行的判据：必须有合法账期 YYYY-MM。
        # 仅看"客户列非空"不够——表尾说明区会在 C 列写长文本，会被误当成客户。
        if not re.fullmatch(r"\d{4}-\d{2}", period):
            continue
        if customer_filter and customer_filter not in customer:
            continue

        pay_day = to_num(cached("pay_day"))
        q = to_num(cached("amount"))
        r_collected = to_num(cached("collected"))
        m_recv = norm(cached("received"))
        n_priced = norm(cached("priced"))
        o_issued = norm(cached("issued"))
        p_orig = norm(cached("original"))

        # 截止日链：有缓存用缓存，否则按公式算（工作日口径由 --calendar 决定）
        f_date = to_date(cached("submit_due")) or period_to_next_month_first(period)
        g_date = to_date(cached("issue_due")) or (add_workdays(f_date, 5, calendar) if f_date else None)
        h_date = to_date(cached("confirm_due")) or (add_workdays(g_date, 5, calendar) if g_date else None)
        i_date = to_date(cached("collect_due")) or next_month_day(period, pay_day)

        # 当前应做 / 距截止 / 紧急度
        todo_cached = norm(cached("todo"))
        if todo_cached:
            todo, todo_src = todo_cached, "缓存"
        else:
            todo = compute_todo(q, r_collected, i_date, p_orig, o_issued, n_priced, m_recv, f_date, today)
            todo_src = "即时算"

        dl_cached = cached("days_left")
        if isinstance(dl_cached, (int, float)):
            days_left, dl_src = int(dl_cached), "缓存"
        else:
            days_left = compute_days_left(todo, i_date, h_date, g_date, f_date, today)
            dl_src = "即时算"

        urg_cached = norm(cached("urgency"))
        urgency = urg_cached or compute_urgency(todo, days_left, q, i_date, r_collected, today)

        items.append({
            "row": r,
            "id": norm(cached("id")),
            "period": period,
            "customer": customer,
            "contract": norm(cached("contract")),
            "pay_day": int(pay_day) if pay_day else None,
            "submit_due": f_date.isoformat() if f_date else "",
            "issue_due": g_date.isoformat() if g_date else "",
            "confirm_due": h_date.isoformat() if h_date else "",
            "collect_due": i_date.isoformat() if i_date else "",
            "todo": todo,
            "days_left": days_left,
            "urgency": urgency,
            "amount": q,
            "collected": r_collected,
            "blocker": norm(cached("blocker")),
            "note": norm(cached("note")),
            "_src": {"todo": todo_src, "days_left": dl_src},
        })
    return items


# ---------- 主流程 ----------

def main():
    ap = argparse.ArgumentParser(description="台账「对账待办」只读查询")
    ap.add_argument("--ledger", required=True, help="台账 xlsx 路径")
    ap.add_argument("--customer", help="只看某家（模糊匹配）")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--calendar", help="法定节假日与调休日历 JSON；提供后工作日按同一日历口径计算")
    args = ap.parse_args()

    try:
        calendar = workdays.load_calendar(args.calendar)
    except workdays.CalendarError as e:
        print(f"[无法解析] {e}", file=sys.stderr)
        return 1
    workday_basis = (WORKDAY_BASIS_WITH_CALENDAR if calendar else WORKDAY_BASIS_WEEKEND_ONLY)

    try:
        wb_v = openpyxl.load_workbook(args.ledger, data_only=True)
        wb_f = openpyxl.load_workbook(args.ledger, data_only=False)
    except Exception as e:
        print(f"[无法读取] {args.ledger}: {e}", file=sys.stderr)
        return 1

    if SHEET_NAME not in wb_v.sheetnames:
        print(f"[无法解析] 台账中没有「{SHEET_NAME}」工作表。现有：{wb_v.sheetnames}", file=sys.stderr)
        return 1

    ws_v, ws_f = wb_v[SHEET_NAME], wb_f[SHEET_NAME]
    header_row = find_header_row(ws_v)
    if header_row is None:
        print("[无法解析] 找不到表头（需含「账期」与「客户」）。", file=sys.stderr)
        return 1

    cols = map_cols(ws_v, header_row)
    today = date.today()
    items = collect(ws_v, ws_f, header_row, cols, today, args.customer, calendar)

    pending = [i for i in items if i["todo"] and i["todo"] != "已完成"]
    overdue = [i for i in pending if i["urgency"] == "逾期"]
    computed = any(i["_src"]["todo"] == "即时算" for i in items)

    result = {
        "ledger": args.ledger,
        "sheet": SHEET_NAME,
        "as_of": today.isoformat(),
        "value_source": "部分列为公式，已按台账规则即时计算（文件无缓存）" if computed else "Excel 缓存",
        "workday_basis": workday_basis,
        "stats": {"total": len(items), "pending": len(pending), "overdue": len(overdue)},
        "items": items,
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    scope = f"（只看「{args.customer}」）" if args.customer else ""
    print(f"台账待办{scope}　基准日 {today.isoformat()}")
    print(f"共 {len(items)} 行：待处理 {len(pending)}，逾期 {len(overdue)}")
    if computed:
        print("注：本文件的公式列无缓存值，以下「当前应做/距截止」按台账规则即时计算。")
    print(f"工作日口径：{workday_basis}。")
    print()

    if not items:
        print("没有符合条件的待办行。")
        return 0

    def sort_key(i):
        d = i["days_left"]
        if d is None:
            return (2, 9999)
        return (0, d) if d < 0 else (1, d)

    for it in sorted(items, key=sort_key):
        if it["todo"] == "已完成" and not it["amount"]:
            continue
        print(f"· {it['customer'] or '（无客户）'}　{it['period'] or '（无账期）'}　→ {it['todo'] or '—'}")
        bits = []
        if it["days_left"] is not None:
            bits.append(f"距截止 {it['days_left']} 天")
        if it["urgency"]:
            bits.append(it["urgency"])
        if it["amount"] is not None:
            bits.append(f"本期结算 {it['amount']:.2f} 元")
        if it["amount"] and it["collected"] is not None:
            bits.append(f"已回款 {it['collected']:.2f} 元")
        if it["collect_due"]:
            bits.append(f"回款截止 {it['collect_due']}")
        if it["blocker"]:
            bits.append(f"卡点 {it['blocker']}")
        if bits:
            print("   " + "　".join(bits))
        print()

    return 0


if __name__ == "__main__":
    sys.exit(main())
