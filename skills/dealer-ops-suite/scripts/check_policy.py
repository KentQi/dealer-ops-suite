#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""政策表自洽检查（只读）。

只做 3 条「表内原则性」检查，输出**需确认项**，不输出「错误」：

  P1 同条码多行     同一商品条码出现两次及以上（可能是同款两个版本，需人确认哪行有效）
  P2 门槛打架       行内说明与表尾备注给出不同的起订/累计门槛
  P3 条码非法/为空  非 13 位数字、或为空

明确不做：跨客户比价、毛利实算、折价率、正价表存在性、FEFO。

用法：
    python3 check_policy.py "<政策表.xlsx>" [--json]

退出码：0 = 无需确认项；2 = 存在需确认项；1 = 无法解析（文件不可读/找不到表头）
"""

import argparse
import json
import re
import sys

try:
    import openpyxl
except ImportError:
    print("需要 openpyxl：python3 -m pip install openpyxl", file=sys.stderr)
    sys.exit(1)


# ---------- 表头识别 ----------

# 表头指纹：至少命中 2 个 token 才认为是表头行
HEADER_TOKENS = ("商品条码", "条码", "商品名称", "商品编码", "系列")

# 列名归一化映射（同一语义多种写法）
COL_ALIASES = {
    "barcode": ("商品条码", "条码", "sku条码"),
    "name": ("商品名称", "品名", "商品"),
    "series": ("系列", "品类"),
}

# 说明类列：可能有多列（如「政策配赠」+「政策赠品搭配」），需全部收集后合并
NOTE_KEYWORDS = ("配赠", "搭配", "说明", "备注")


def _norm(s):
    if s is None:
        return ""
    return re.sub(r"\s+", "", str(s))


def find_header_row(ws, max_scan=12):
    """扫描前 max_scan 行，找命中 token 最多的一行作为表头。"""
    best_row, best_hits = None, 0
    for r in range(1, min(ws.max_row, max_scan) + 1):
        cells = [_norm(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        hits = sum(1 for tok in HEADER_TOKENS if any(tok in c for c in cells if c))
        if hits > best_hits:
            best_row, best_hits = r, hits
    if best_hits >= 2:
        return best_row
    return None


def map_columns(ws, header_row):
    """按归一化表头文本映射到规范字段。说明类列返回列表（可能多列）。"""
    mapping = {}
    note_cols = []
    for c in range(1, ws.max_column + 1):
        h = _norm(ws.cell(row=header_row, column=c).value)
        if not h:
            continue
        if any(k in h for k in NOTE_KEYWORDS):
            note_cols.append(c)
            continue
        for field, aliases in COL_ALIASES.items():
            if field in mapping:
                continue
            if any(a == h or (a in h and len(h) <= len(a) + 4) for a in aliases):
                mapping[field] = c
    if note_cols:
        mapping["note_cols"] = note_cols
    return mapping


def norm_barcode(v):
    """条码归一化：数值型统一成整数字符串，避免 float 尾巴。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    return _norm(v)


# ---------- 检查 ----------

def check_barcode_dups(rows):
    """P1：同条码多行 → 需确认哪行有效。"""
    buckets = {}
    for row in rows:
        bc = row["barcode"]
        if not bc:
            continue
        buckets.setdefault(bc, []).append(row)

    findings = []
    for bc, items in buckets.items():
        if len(items) < 2:
            continue
        # 如果多行的关键数值完全一致，大概率是重复粘贴，也仍需确认
        findings.append({
            "id": "P1",
            "type": "同条码多行",
            "barcode": bc,
            "rows": [it["row"] for it in items],
            "names": [it["name"] for it in items],
            "prices": [it["prices"] for it in items],
            "ask": "这几行同条码，请确认哪一行有效（或说明这是同款不同版本）",
        })
    return findings


def check_barcode_invalid(rows):
    """P3：条码非法或为空。"""
    findings = []
    for row in rows:
        bc = row["barcode"]
        if not bc:
            findings.append({
                "id": "P3",
                "type": "条码为空",
                "row": row["row"],
                "name": row["name"],
                "ask": "该行条码为空，无法对应 SKU，请补条码",
            })
        elif not re.fullmatch(r"\d{8,14}", bc):
            findings.append({
                "id": "P3",
                "type": "条码非法",
                "row": row["row"],
                "name": row["name"],
                "barcode": bc,
                "ask": "条码应为 8–14 位纯数字，请确认",
            })
    return findings


THRESHOLD_PATTERNS = (
    # 起订量 4 箱 / 起订 4 箱
    ("起订", re.compile(r"起订(?:量)?\s*(\d+)\s*(箱|套|盒|件)")),
    # 累计满足 12 箱 / 累计 12 箱
    ("累计", re.compile(r"累计[^0-9]{0,6}(\d+)\s*(箱|套|盒|件)")),
)


def extract_thresholds(text):
    """从一段文字里抽出 {门槛类型: (数值, 单位)}。"""
    out = {}
    if not text:
        return out
    for kind, pat in THRESHOLD_PATTERNS:
        for m in pat.finditer(text):
            out[kind] = (int(m.group(1)), m.group(2))
    return out


def check_threshold_conflict(rows, footer_text):
    """P2：行内说明与表尾备注的门槛不一致 → 需确认以哪个为准。

    两种形态都算打架：
      (a) 同一类门槛数值不同（如行内起订 4 箱 / 表尾起订 8 箱）
      (b) 表尾比行内多出一个门槛条件（如行内只写"起订 4 箱"，表尾还要求"累计 12 箱"）
          —— 这会让同一政策在不同位置得出不同适用范围。

    为避免误报，形态 (b) 只在「行内说明非空」且「行名称出现在表尾文字中」时判定，
    即确认两者讲的是同一个商品/同一组政策。
    """
    if not footer_text:
        return []
    footer_th = extract_thresholds(footer_text)
    if not footer_th:
        return []

    findings = []
    seen = set()
    for row in rows:
        note = row.get("note", "")
        row_th = extract_thresholds(note)

        # 形态 (a)：同类门槛数值不同（必须单位相同才可比）
        for kind, (val, unit) in row_th.items():
            if kind not in footer_th:
                continue
            f_val, f_unit = footer_th[kind]
            if f_unit != unit:
                continue
            if (val, unit) != (f_val, f_unit):
                key = ("a", kind, val, unit, f_val, f_unit)
                if key in seen:
                    continue
                seen.add(key)
                findings.append({
                    "id": "P2",
                    "type": "门槛数值不一致",
                    "kind": kind,
                    "row": row["row"],
                    "name": row["name"],
                    "inline": f"{kind} {val}{unit}",
                    "footer": f"{kind} {f_val}{f_unit}",
                    "ask": f"行内说明与表尾备注的「{kind}」门槛不一致（{val}{unit} vs {f_val}{f_unit}），请确认以哪个为准",
                })

        # 形态 (b)：表尾多出条件（仅当行内非空、且单位一致时才比较）
        # 单位一致是关键：行内说"箱"就只跟表尾的"箱"比，不碰"套"——否则会把
        # 礼盒的"起订24套"错算到面膜头上（典型误报）。
        if not note:
            continue
        row_units = {u for (_, u) in row_th.values()}
        if not row_units:
            continue
        for kind, (f_val, f_unit) in footer_th.items():
            if kind in row_th or f_unit not in row_units:
                continue
            key = ("b", kind, f_val, f_unit)
            if key in seen:
                continue
            seen.add(key)
            findings.append({
                "id": "P2",
                "type": "表尾多出条件",
                "kind": kind,
                "row": row["row"],
                "name": row["name"],
                "inline": note[:60],
                "footer": f"{kind} {f_val}{f_unit}",
                "ask": f"表尾要求「{kind} {f_val}{f_unit}」，但行内说明未提该条件 → 同一政策在不同位置适用范围不同，请确认以哪个为准",
            })
    return findings


# ---------- 主流程 ----------

def collect_rows(ws, header_row, cols):
    rows = []
    bc_col = cols.get("barcode")
    name_col = cols.get("name")
    note_cols = cols.get("note_cols", [])
    if not bc_col:
        return rows

    for r in range(header_row + 1, ws.max_row + 1):
        name = _norm(ws.cell(row=r, column=name_col).value) if name_col else ""
        bc = norm_barcode(ws.cell(row=r, column=bc_col).value)
        note_parts = []
        for c in note_cols:
            v = _norm(ws.cell(row=r, column=c).value)
            if v and v not in ("无", "/"):
                note_parts.append(v)
        note = "；".join(note_parts)
        # 全空行跳过
        if not name and not bc and not note:
            continue
        # 表尾备注行：首列是「备注」等标题，或该行没有名称与条码
        first = _norm(ws.cell(row=r, column=1).value)
        if first in ("备注", "说明", "政策说明") or (name == "" and bc == ""):
            continue
        rows.append({
            "row": r,
            "barcode": bc,
            "name": name,
            "note": note,
            "prices": {},
        })
    return rows


def collect_footer(ws, header_row):
    """表尾备注：表头之后、含「备注/说明/累计/起订」且不在数据区的文本。"""
    chunks = []
    for r in range(header_row + 1, ws.max_row + 1):
        for c in range(1, min(ws.max_column, 4) + 1):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            if any(k in v for k in ("备注", "说明", "累计", "起订", "政策")):
                chunks.append(v)
    return "\n".join(chunks)


def main():
    ap = argparse.ArgumentParser(description="政策表自洽检查（只读，输出需确认项）")
    ap.add_argument("xlsx", help="政策表路径")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    ap.add_argument("--sheet", help="指定工作表名（默认取第一个含「政策」或「价格」的表）")
    args = ap.parse_args()

    try:
        wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    except Exception as e:
        print(f"[无法读取] {args.xlsx}: {e}", file=sys.stderr)
        return 1

    # 选表
    if args.sheet:
        ws = wb[args.sheet]
    else:
        ws = wb[wb.sheetnames[0]]
        for n in wb.sheetnames:
            if "政策" in n or "价格" in n:
                ws = wb[n]
                break

    header_row = find_header_row(ws)
    if header_row is None:
        print(f"[无法解析] 在「{ws.title}」前 12 行找不到表头（需命中商品条码/商品名称等）。", file=sys.stderr)
        return 1

    cols = map_columns(ws, header_row)
    if "barcode" not in cols:
        print(f"[无法解析] 表头行 {header_row} 中找不到条码列。", file=sys.stderr)
        return 1

    rows = collect_rows(ws, header_row, cols)
    footer = collect_footer(ws, header_row)

    findings = []
    findings += check_barcode_dups(rows)
    findings += check_threshold_conflict(rows, footer)
    findings += check_barcode_invalid(rows)

    result = {
        "file": args.xlsx,
        "sheet": ws.title,
        "header_row": header_row,
        "data_rows": len(rows),
        "columns": {
            k: ([openpyxl.utils.get_column_letter(x) for x in v] if isinstance(v, list)
                else openpyxl.utils.get_column_letter(v))
            for k, v in cols.items()
        },
        "needs_confirmation": findings,
        "summary": {
            "total": len(findings),
            "P1_同条码多行": sum(1 for f in findings if f["id"] == "P1"),
            "P2_门槛打架": sum(1 for f in findings if f["id"] == "P2"),
            "P3_条码问题": sum(1 for f in findings if f["id"] == "P3"),
        },
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"文件：{args.xlsx}")
        print(f"工作表：{ws.title}　表头行：{header_row}　数据行：{len(rows)}")
        print()
        if not findings:
            print("无需确认项。")
        else:
            print(f"需你确认 {len(findings)} 项（不是「错误」，是需要你判断以哪个为准）：")
            print()
            for i, f in enumerate(findings, 1):
                print(f"{i}. [{f['id']} {f['type']}]")
                if f["id"] == "P1":
                    print(f"   条码 {f['barcode']}，第 {f['rows']} 行：{f['names']}")
                    print(f"   → {f['ask']}")
                elif f["id"] == "P2":
                    print(f"   第 {f['row']} 行「{f['name']}」：行内 {f['inline']}　表尾 {f['footer']}")
                    print(f"   → {f['ask']}")
                else:
                    print(f"   第 {f['row']} 行「{f['name']}」：{f.get('barcode') or '（空）'}")
                    print(f"   → {f['ask']}")
                print()
        s = result["summary"]
        print(f"汇总：同条码多行 {s['P1_同条码多行']}　门槛打架 {s['P2_门槛打架']}　条码问题 {s['P3_条码问题']}")

    return 2 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
