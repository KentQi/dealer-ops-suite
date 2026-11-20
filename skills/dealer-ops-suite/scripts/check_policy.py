#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""价格／出货政策只读审阅。

同条码的连续生效版本不是重复行。缺结算价、非法值、区间重叠属于阻断；
生效断档、条码或门槛等业务歧义进入需确认项。脚本不判断跨客户价格、
毛利或折扣是否合理，也不替用户改写政策。

用法：python3 check_policy.py 政策表.xlsx [--json] [--sheet 工作表]
退出码：0 无问题；2 有阻断或需确认项；1 文件／表头无法解析。
"""
import argparse
import json
import os
import re
import sys
from datetime import timedelta

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from settlement_rules import RateBook, RuleViolation, parse_day, rate_value_column

HEADER_TOKENS = ("商品条码", "条码", "商品名称", "商品编码", "系列")
NOTE_KEYWORDS = ("配赠", "搭配", "说明", "备注")
COL_ALIASES = {
    "barcode": ("商品条码", "条码", "sku条码"),
    "name": ("商品名称", "品名"),
    "price": ("结算价", "合同价", "供货价", "出货价", "供应价"),
    "decal": ("贴花标准", "贴花单价"),
    "start": ("生效起", "生效开始", "生效日期", "开始日期"),
    "end": ("生效止", "生效结束", "失效日期", "结束日期"),
    "source": ("依据文件", "政策编号", "版本号"),
}


class PolicyColumnError(ValueError):
    """价格／标准列存在多个权威候选，必须与对账引擎同样阻断。"""


def _norm(value):
    return "" if value is None else re.sub(r"\s+", "", str(value))


def find_header_row(ws, max_scan=12):
    best_row, best_hits = None, 0
    for row in range(1, min(ws.max_row, max_scan) + 1):
        cells = [_norm(cell.value) for cell in ws[row]]
        hits = sum(1 for token in HEADER_TOKENS if any(token in cell for cell in cells if cell))
        if hits > best_hits:
            best_row, best_hits = row, hits
    return best_row if best_hits >= 2 else None


def map_columns(ws, header_row):
    mapping, notes, assigned = {}, [], {}
    headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
    # 价格／贴花列与两个对账引擎共用同一份解析：多列命中即阻断，不任取左列。
    for field in ("price", "decal"):
        try:
            mapping[field] = rate_value_column(headers, COL_ALIASES[field], f"政策表·{ws.title}") + 1
        except RuleViolation as exc:
            if exc.code == "AMBIGUOUS_RATE_COLUMN":
                raise PolicyColumnError(exc.message) from exc
            # 缺失价格列由 check_versions 按 blocker 输出；贴花列本身可选。
            continue
        assigned[mapping[field]] = field
    for col in range(1, ws.max_column + 1):
        header = _norm(ws.cell(header_row, col).value)
        if not header:
            continue
        if any(word in header for word in NOTE_KEYWORDS):
            notes.append(col)
            continue
        for field, aliases in COL_ALIASES.items():
            if field in mapping:
                continue
            if any(header == alias or (alias in header and len(header) <= len(alias) + 5)
                   for alias in aliases):
                if col in assigned:
                    raise ValueError(f"表头「{header}」同时映射为 {assigned[col]} 与 {field}，请修正表头")
                mapping[field] = col
                assigned[col] = field
    if notes:
        mapping["note_cols"] = notes
    return mapping


def norm_barcode(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(int(value)) if isinstance(value, int) else _norm(value)


def collect_rows(ws, header_row, cols):
    rows = []
    # 数据行判据只看条码／品名／价格／标准／生效区间；仅备注或依据文件有文字的表尾说明行继续豁免。
    core_cols = [cols[field] for field in ("barcode", "name", "price", "decal", "start", "end") if field in cols]
    for row in range(header_row + 1, ws.max_row + 1):
        if not any(_norm(ws.cell(row, col).value) for col in core_cols):
            continue
        barcode = norm_barcode(ws.cell(row, cols["barcode"]).value)
        name = _norm(ws.cell(row, cols["name"]).value) if "name" in cols else ""
        notes = [_norm(ws.cell(row, col).value) for col in cols.get("note_cols", [])]
        note = "；".join(value for value in notes if value and value not in ("无", "/"))
        item = {"row": row, "barcode": barcode, "name": name, "note": note}
        for field in ("price", "decal", "start", "end", "source"):
            item[field] = ws.cell(row, cols[field]).value if field in cols else None
        rows.append(item)
    return rows


def collect_footer(ws, header_row):
    chunks = []
    for row in range(header_row + 1, ws.max_row + 1):
        for col in range(1, min(ws.max_column, 4) + 1):
            value = ws.cell(row, col).value
            if isinstance(value, str) and any(word in value for word in ("备注", "说明", "累计", "起订", "政策")):
                chunks.append(value)
    return "\n".join(chunks)


THRESHOLD_PATTERNS = (
    ("起订", re.compile(r"起订(?:量)?\s*(\d+)\s*(箱|套|盒|件)")),
    ("累计", re.compile(r"累计[^0-9]{0,6}(\d+)\s*(箱|套|盒|件)")),
)


def extract_thresholds(text):
    return {kind: (int(match.group(1)), match.group(2))
            for kind, pattern in THRESHOLD_PATTERNS for match in pattern.finditer(text or "")}


def finding(code, kind, message, severity="high", **context):
    return {"id": code, "type": kind, "severity": severity, "ask": message,
            "message": message, **context}


def check_barcodes(rows):
    output = []
    for row in rows:
        barcode = row["barcode"]
        if not barcode:
            output.append(finding("P3", "条码为空", "该行条码为空，无法关联 SKU，请补齐。", "blocker",
                                  row=row["row"], name=row["name"]))
        elif not re.fullmatch(r"\d{8,14}", barcode):
            output.append(finding("P3", "条码非法", "条码须为 8–14 位纯数字，请修正。", "blocker",
                                  row=row["row"], barcode=barcode, name=row["name"]))
    return output


def check_threshold_conflict(rows, footer):
    footer_thresholds = extract_thresholds(footer)
    if not footer_thresholds:
        return []
    output, seen = [], set()
    for row in rows:
        inline = extract_thresholds(row["note"])
        units = {unit for _, unit in inline.values()}
        for kind, (value, unit) in inline.items():
            if kind in footer_thresholds and footer_thresholds[kind] != (value, unit) and footer_thresholds[kind][1] == unit:
                other = footer_thresholds[kind][0]
                key = (kind, value, other, unit)
                if key not in seen:
                    seen.add(key)
                    output.append(finding("P2", "门槛数值不一致", "行内说明与表尾门槛冲突，请确认适用范围。",
                                          row=row["row"], name=row["name"], kind=kind,
                                          inline=f"{value}{unit}", footer=f"{other}{unit}"))
        if not row["note"] or not units:
            continue
        for kind, (value, unit) in footer_thresholds.items():
            key = (kind, value, unit)
            if kind not in inline and unit in units and key not in seen:
                seen.add(key)
                output.append(finding("P2", "表尾多出条件", "行内与表尾适用条件不同，请确认以哪个为准。",
                                      row=row["row"], name=row["name"], kind=kind,
                                      inline=row["note"][:80], footer=f"{value}{unit}"))
    return output


def check_versions(rows, cols):
    """共享 RateBook 的区间语义；重叠阻断，未来／相邻版本断档需确认。"""
    output = []
    if "price" not in cols:
        return [finding("P4", "缺少结算价列", "最终价格政策必须提供结算价／合同价／供货价列。", "blocker")]
    for field, label in (("start", "生效起"), ("price", "结算价")):
        if field not in cols:
            output.append(finding("P4", f"缺少{label}列", f"价格政策缺少{label}列，无法确定结算取价。", "blocker"))
    if output:
        return output
    for row in rows:
        if row["price"] in (None, ""):
            output.append(finding("P4", "结算价为空", "该行没有结算价，不得按其他版本猜价。", "blocker",
                                  row=row["row"], barcode=row["barcode"]))
        if row["start"] in (None, ""):
            output.append(finding("P4", "生效日为空", "价格版本必须有明确生效日。", "blocker",
                                  row=row["row"], barcode=row["barcode"]))
    rate_input = [{"barcode": row["barcode"], "value": row["price"], "start": row["start"],
                   "end": row["end"], "source": row["source"] or f"第 {row['row']} 行"}
                  for row in rows if row["barcode"] and row["price"] not in (None, "") and row["start"] not in (None, "")]
    try:
        book = RateBook(rate_input, kind="结算价")
    except RuleViolation as exc:
        output.append(finding("P4", "价格版本原则性错误", exc.message, "blocker", rule_code=exc.code,
                              context=exc.context))
        book = None
    if book is not None:
        for barcode, rates in book._rates.items():
            for previous, current in zip(rates, rates[1:]):
                if previous.end is not None and previous.end + timedelta(days=1) < current.start:
                    output.append(finding("P5", "价格生效断档", "相邻价格版本之间有未覆盖日期，请补充生效区间或说明不适用日期。",
                                          barcode=barcode, gap_start=(previous.end + timedelta(days=1)).isoformat(),
                                          gap_end=(current.start - timedelta(days=1)).isoformat()))
    if "decal" in cols:
        decal_input = [{"barcode": row["barcode"], "value": row["decal"], "start": row["start"],
                        "end": row["end"], "source": row["source"] or f"第 {row['row']} 行"}
                       for row in rows if row["barcode"] and row["decal"] not in (None, "") and row["start"] not in (None, "")]
        try:
            RateBook(decal_input, kind="贴花标准")
        except RuleViolation as exc:
            output.append(finding("P6", "贴花标准原则性错误", exc.message, "blocker", rule_code=exc.code,
                                  context=exc.context))
    # 同码多行若不能组成带生效日的版本序列，才作为业务歧义提示。
    buckets = {}
    for row in rows:
        if row["barcode"]:
            buckets.setdefault(row["barcode"], []).append(row)
    for barcode, items in buckets.items():
        if len(items) > 1 and any(item["start"] in (None, "") for item in items):
            output.append(finding("P1", "同条码多行", "同码多行但未逐行标生效日，请确认各版本生效范围。", barcode=barcode,
                                  rows=[item["row"] for item in items]))
    return output


def analyse(xlsx, sheet=None):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    try:
        if sheet:
            if sheet not in wb.sheetnames:
                raise ValueError("工作表不存在：" + sheet)
            ws = wb[sheet]
        else:
            ws = next((wb[name] for name in wb.sheetnames if "政策" in name or "价格" in name), wb[wb.sheetnames[0]])
        header_row = find_header_row(ws)
        if header_row is None:
            raise ValueError("前 12 行找不到商品条码／商品名称表头")
        try:
            cols = map_columns(ws, header_row)
        except PolicyColumnError as exc:
            blocker = finding("P4", "价格／标准列歧义", str(exc), "blocker")
            return {"file": os.path.abspath(xlsx), "sheet": ws.title, "header_row": header_row,
                    "data_rows": 0, "columns": {},
                    "needs_confirmation": [blocker], "blocking_findings": [blocker],
                    "summary": {"total": 1, "blocker": 1, "high": 0}}
        if "barcode" not in cols:
            raise ValueError("找不到条码列")
        rows = collect_rows(ws, header_row, cols)
        if not rows:
            raise ValueError("未读取到 SKU 政策行")
        findings = (check_barcodes(rows) + check_threshold_conflict(rows, collect_footer(ws, header_row))
                    + check_versions(rows, cols))
        return {"file": os.path.abspath(xlsx), "sheet": ws.title, "header_row": header_row,
                "data_rows": len(rows), "columns": {key: value for key, value in cols.items()},
                "needs_confirmation": findings,
                "blocking_findings": [item for item in findings if item["severity"] == "blocker"],
                "summary": {"total": len(findings),
                            "blocker": sum(item["severity"] == "blocker" for item in findings),
                            "high": sum(item["severity"] == "high" for item in findings)}}
    finally:
        wb.close()


def main():
    parser = argparse.ArgumentParser(description="最终价格／出货政策只读审阅")
    parser.add_argument("xlsx")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--sheet")
    args = parser.parse_args()
    try:
        result = analyse(args.xlsx, args.sheet)
    except Exception as exc:
        print("[无法解析] " + str(exc), file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"文件：{result['file']}\n工作表：{result['sheet']}　数据行：{result['data_rows']}")
        if not result["needs_confirmation"]:
            print("未发现原则性错误或待确认项。")
        for item in result["needs_confirmation"]:
            print(f"[{item['severity']} {item['id']} {item['type']}] {item['ask']}")
        print(f"汇总：阻断 {result['summary']['blocker']}　需确认 {result['summary']['high']}")
    return 2 if result["needs_confirmation"] else 0


if __name__ == "__main__":
    sys.exit(main())
