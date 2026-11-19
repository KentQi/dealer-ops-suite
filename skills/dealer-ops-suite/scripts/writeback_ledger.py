#!/usr/bin/env python3
"""把对账结果回写主档「期间」表（B5）。

默认 dry-run：只打印将要发生的变更，不写文件。
加 --write 才落盘；写前自动备份，写后原子替换。

幂等：同一 period 已存在则更新该行，不追加重复行。

用法：
    python3 writeback_ledger.py "<_主档.xlsx>" --period 202610 \
        --asof 2026-11-02 --opening-qty 1000 --opening-amount 50000 \
        --settle-qty 120 --settle-amount 5000 --diff-qty -20 --diff-amount -800 \
        --sticker-amount 200 --sticker-pay "待定" --status "待确认"

    # 从 JSON 读取（键可用英文或中文）
    python3 writeback_ledger.py "<_主档.xlsx>" --from-json /tmp/writeback.json

    # 确认后落盘
    python3 writeback_ledger.py "<_主档.xlsx>" --from-json /tmp/writeback.json --write

JSON 示例：
{
  "period": "202610",
  "asof": "2026-11-02",
  "opening_qty": 1000,
  "opening_amount": 50000,
  "settle_qty": 120,
  "settle_amount": 5000,
  "diff_qty": -20,
  "diff_amount": -800,
  "sticker_amount": 200,
  "sticker_pay": "待定",
  "statement_no": "",
  "status": "待确认",
  "note": ""
}
"""
import argparse
import json
import math
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from copy import copy
from datetime import datetime

import openpyxl

# key -> (中文表头, 允许的别名)
FIELDS = [
    ("period", "期间", ("期间", "对账期间", "月份")),
    ("asof", "数据基准日", ("数据基准日", "基准日", "库存基准日")),
    ("opening_qty", "期初数量", ("期初数量", "期初件数")),
    ("opening_amount", "期初金额", ("期初金额", "期初货值")),
    ("settle_qty", "结算数量", ("结算数量", "结算件数")),
    ("settle_amount", "结算金额", ("结算金额",)),
    ("diff_qty", "差异数量", ("差异数量", "差异件数")),
    ("diff_amount", "差异金额", ("差异金额",)),
    ("sticker_amount", "贴花金额", ("贴花金额", "贴花费用")),
    ("sticker_pay", "贴花兑付方式", ("贴花兑付方式", "兑付方式", "贴花兑付")),
    ("statement_no", "对账单编号", ("对账单编号", "对账单号")),
    ("status", "状态", ("状态",)),
    ("note", "备注", ("备注",)),
]

NUMERIC = {
    "opening_qty", "opening_amount", "settle_qty", "settle_amount",
    "diff_qty", "diff_amount", "sticker_amount",
}
NON_NEGATIVE = {
    "opening_qty", "opening_amount", "settle_qty", "settle_amount",
    "sticker_amount",
}
STATUSES = {"待确认", "已确认", "已对账", "已付款", "已签回", "已收款", "待处理", "已作废"}
PAY_MODES = {"抵扣货款", "分开支付", "待定", ""}


def norm(value):
    """归一化表头/枚举：全角转半角、去空白、去括号说明。"""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip()
    text = re.sub(r"\s+", "", text)
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\(.*?\)", "", text)
    return text.lower()


def fail(message, code=1):
    print(f"[中止] {message}", file=sys.stderr)
    sys.exit(code)


def parse_number(value, key, label):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        fail(f"{label} 不是数字：{value!r}")
    try:
        number = float(value)
    except (TypeError, ValueError):
        fail(f"{label} 不是数字：{value!r}")
    if not math.isfinite(number):
        fail(f"{label} 不是有限数字：{value!r}")
    if abs(number) >= 1e12:
        fail(f"{label} 超出允许范围（|x| < 1e12）：{value!r}")
    if key in NON_NEGATIVE and number < 0:
        fail(f"{label} 不能为负数：{value!r}")
    return round(number, 2)


def load_payload(args):
    payload = {}
    if args.from_json:
        try:
            with open(args.from_json, encoding="utf-8") as fh:
                payload = json.load(fh)
        except Exception as exc:
            fail(f"读取 JSON 失败：{exc}")
        if not isinstance(payload, dict):
            fail("JSON 顶层必须是对象")
    return payload


def pick(args, payload, key, label):
    """命令行优先，其次 JSON 的英文键 / 中文键。"""
    cli = getattr(args, key, None)
    if cli is not None:
        return cli
    if key in payload:
        return payload[key]
    if label in payload:
        return payload[label]
    return None


def collect(args):
    payload = load_payload(args)
    values = {}
    for key, label, _aliases in FIELDS:
        values[key] = pick(args, payload, key, label)

    if not values["period"]:
        fail("缺少 --period（YYYYMM）或 JSON 的 period/期间")
    values["period"] = str(values["period"]).strip()
    if not re.fullmatch(r"\d{6}", values["period"]):
        fail(f"期间必须是 6 位年月，如 202610；当前 {values['period']!r}")
    if not 1 <= int(values["period"][4:]) <= 12:
        fail(f"期间月份非法：{values['period']!r}")

    if values["asof"]:
        values["asof"] = str(values["asof"]).strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", values["asof"]):
            fail(f"数据基准日必须是 YYYY-MM-DD；当前 {values['asof']!r}")

    for key in NUMERIC:
        values[key] = parse_number(values[key], key, dict((k, l) for k, l, _ in FIELDS)[key])

    for key in ("sticker_pay", "status", "statement_no", "note"):
        value = values[key]
        values[key] = None if value is None else str(value).strip()

    if values["sticker_pay"] is not None and values["sticker_pay"] not in PAY_MODES:
        print(f"[提示] 贴花兑付方式不在常见取值内：{values['sticker_pay']!r}", file=sys.stderr)
    if values["status"] and values["status"] not in STATUSES:
        print(f"[提示] 状态不在常见取值内：{values['status']!r}", file=sys.stderr)

    if not any(values[k] is not None for k, _, _ in FIELDS if k != "period"):
        fail("除期间外没有任何要写入的字段")
    return values


def header_map(headers):
    mapping = {}
    for index, header in enumerate(headers, 1):
        head = norm(header)
        if not head:
            continue
        for key, _label, aliases in FIELDS:
            if key in mapping:
                continue
            if head in {norm(alias) for alias in aliases}:
                mapping[key] = index
    return mapping


def locate_period_sheet(wb):
    candidates = []
    for ws in wb.worksheets:
        max_scan = min(ws.max_row, 12)
        for row in range(1, max_scan + 1):
            headers = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]
            mapping = header_map(headers)
            if "period" in mapping and ("settle_amount" in mapping or "diff_amount" in mapping):
                candidates.append((ws, row, mapping))
                break
    if not candidates:
        fail("未找到「期间」表：需要同时有「期间」和「结算金额/差异金额」表头", 2)
    if len(candidates) > 1:
        names = "、".join(f"{ws.title}!{row}" for ws, row, _ in candidates)
        fail(f"找到多个可能的「期间」表：{names}；请先整理主档", 2)
    return candidates[0]


def last_data_row(ws, header_row, mapping):
    last = header_row
    cols = sorted(set(mapping.values()))
    for row in range(header_row + 1, ws.max_row + 1):
        if any(ws.cell(row, col).value not in (None, "") for col in cols):
            last = row
    return last


def find_period_row(ws, header_row, mapping, period):
    col = mapping["period"]
    rows = []
    for row in range(header_row + 1, ws.max_row + 1):
        if norm(ws.cell(row, col).value) == norm(period):
            rows.append(row)
    if len(rows) > 1:
        fail(f"期间 {period} 在主档中出现多行：{rows}；请先合并", 2)
    return rows[0] if rows else None


def format_change(label, old, new):
    def show(value):
        if value is None or value == "":
            return "（空）"
        if isinstance(value, float):
            return f"{value:,.2f}"
        return str(value)
    return f"  {label}: {show(old)} -> {show(new)}"


def print_table(ws, row, mapping, changes):
    print(f"目标行：第 {row} 行")
    for key, label, _aliases in FIELDS:
        if key not in changes:
            continue
        old, new = changes[key]
        print(format_change(label, old, new))


def apply_changes(ws, row, mapping, changes):
    for key, (_old, new) in changes.items():
        ws.cell(row, mapping[key]).value = new


def copy_row_style(ws, source_row, target_row, mapping):
    if source_row < 1:
        return
    if ws.row_dimensions[source_row].height:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in sorted(set(mapping.values())):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        try:
            target._style = copy(source._style)
        except Exception:
            pass


def backup_path(ledger):
    stem, ext = os.path.splitext(ledger)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{stem}.bak-{stamp}{ext}"


def save_atomic(wb, ledger):
    directory = os.path.dirname(os.path.abspath(ledger))
    fd, temp_path = tempfile.mkstemp(prefix=".writeback-", suffix=".xlsx", dir=directory)
    os.close(fd)
    try:
        wb.save(temp_path)
        os.replace(temp_path, ledger)
    except Exception:
        if os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except OSError:
                pass
        raise


def main():
    ap = argparse.ArgumentParser(description="把对账结果回写主档「期间」表（默认 dry-run）")
    ap.add_argument("ledger", help="_主档.xlsx 路径")
    ap.add_argument("--period", help="YYYYMM，如 202610")
    ap.add_argument("--asof", help="数据基准日 YYYY-MM-DD")
    ap.add_argument("--opening-qty", type=float, help="期初数量")
    ap.add_argument("--opening-amount", type=float, help="期初金额")
    ap.add_argument("--settle-qty", type=float, help="结算数量")
    ap.add_argument("--settle-amount", type=float, help="结算金额")
    ap.add_argument("--diff-qty", type=float, help="差异数量（可负）")
    ap.add_argument("--diff-amount", type=float, help="差异金额（可负）")
    ap.add_argument("--sticker-amount", type=float, help="贴花金额")
    ap.add_argument("--sticker-pay", help="贴花兑付方式：抵扣货款/分开支付/待定")
    ap.add_argument("--statement-no", help="对账单编号")
    ap.add_argument("--status", help="状态：待确认/已确认/已对账/已付款/已签回/已收款")
    ap.add_argument("--note", help="备注")
    ap.add_argument("--from-json", help="从 JSON 读取字段（英文或中文键）")
    ap.add_argument("--write", action="store_true", help="确认写入；不加则只 dry-run")
    ap.add_argument("--json", action="store_true", help="输出机器可读结果")
    args = ap.parse_args()

    ledger = os.path.abspath(args.ledger)
    if not os.path.isfile(ledger):
        fail(f"主档不存在：{ledger}")

    values = collect(args)

    try:
        wb = openpyxl.load_workbook(ledger, keep_vba=ledger.lower().endswith(".xlsm"))
    except Exception as exc:
        fail(f"打开主档失败：{exc}")

    ws, header_row, mapping = locate_period_sheet(wb)
    row = find_period_row(ws, header_row, mapping, values["period"])
    created = False
    if row is None:
        previous = last_data_row(ws, header_row, mapping)
        row = previous + 1
        created = True
        copy_row_style(ws, previous, row, mapping)
        ws.cell(row, mapping["period"]).value = values["period"]

    changes = {}
    for key, label, _aliases in FIELDS:
        new = values[key]
        if new is None:
            continue
        if key not in mapping:
            print(f"[提示] 主档「期间」表没有「{label}」列，已跳过", file=sys.stderr)
            continue
        old = ws.cell(row, mapping[key]).value
        if norm(old) == norm(new) and (old in (None, "") or new in (None, "")):
            continue
        if old == new:
            continue
        changes[key] = (old, new)

    if not changes:
        result = {"ok": True, "mode": "noop", "ledger": ledger, "period": values["period"], "row": row}
        print(f"期间 {values['period']} 无变化，未写入。")
        if args.json:
            print(json.dumps(result, ensure_ascii=False))
        return

    print_table(ws, row, mapping, changes)
    if created:
        print(f"  （期间 {values['period']} 不存在，将追加新行）")

    if not args.write:
        print("\n[dry-run] 未写入。确认后加 --write。")
        if args.json:
            print(json.dumps({
                "ok": True, "mode": "dry-run", "ledger": ledger,
                "period": values["period"], "row": row, "created": created,
                "changes": {k: {"old": v[0], "new": v[1]} for k, v in changes.items()},
            }, ensure_ascii=False, default=str))
        return

    backup = backup_path(ledger)
    try:
        shutil.copy2(ledger, backup)
    except Exception as exc:
        fail(f"备份失败，未写入：{exc}", 3)

    apply_changes(ws, row, mapping, changes)
    try:
        save_atomic(wb, ledger)
    except Exception as exc:
        fail(f"写入失败，原文件未替换；备份在：{backup}\n{exc}", 3)

    result = {
        "ok": True, "mode": "write", "ledger": ledger, "period": values["period"],
        "row": row, "created": created, "backup": backup,
        "changes": {k: {"old": v[0], "new": v[1]} for k, v in changes.items()},
    }
    print(f"\n[已写入] 期间 {values['period']} → 第 {row} 行")
    print(f"备份：{backup}")
    if args.json:
        print(json.dumps(result, ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()
