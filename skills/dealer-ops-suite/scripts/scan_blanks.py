#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scan_blanks.py —— 合同模板「待填项」扫描器（只读）

用途：
  1. 阶段 0 识别：按 spec 定位每一个待填项，判断当前是空白 / 已填 / 模板漂移；
  2. 阶段 4 验收：填充后重跑，确认无残留占位符、已填值与计划一致；
  3. 越界预警：列出文档中所有「未被登记」的空白痕迹，防止漏填或模板改版。

本脚本**只读**，绝不修改任何文件。

用法：
  python3 scan_blanks.py --spec spec/jx.json --doc 合同.docx
  python3 scan_blanks.py --spec spec/jx.json --doc 合同.docx --json-out scan.json
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("缺少依赖：pip install python-docx")

# 未被登记时用来找「疑似空白」的模式
BLANK_PATTERNS = [
    r"＿{2,}",
    r"(?<![A-Za-z0-9])_{3,}",
    r"【\s*】",
    r"\(\s{2,}\)",
    r"（\s{2,}）",
    r"X{3,}",
    r"x{3,}",
]


def iter_blocks(doc):
    """按文档真实顺序产出 ('p', Paragraph) / ('tbl', Table, table_index)"""
    ti = -1
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield "p", Paragraph(child, doc), None
        elif child.tag == qn("w:tbl"):
            ti += 1
            yield "tbl", Table(child, doc), ti


def build_units(doc):
    """把正文段落与表格单元格段落拍平成有序文本单元，用于文档级出现次数统计。"""
    units = []
    para_idx = -1
    for kind, blk, ti in iter_blocks(doc):
        if kind == "p":
            para_idx += 1
            units.append(
                {"kind": "p", "para_index": para_idx, "table": None, "row": None,
                 "col": None, "text": blk.text}
            )
        else:
            seen = set()
            for ri, row in enumerate(blk.rows):
                for ci, cell in enumerate(row.cells):
                    key = cell._tc
                    if key in seen:
                        continue
                    seen.add(key)
                    for p in cell.paragraphs:
                        units.append(
                            {"kind": "cell", "para_index": None, "table": ti,
                             "row": ri, "col": ci, "text": p.text}
                        )
    return units


def count_occurrences(units, needle):
    """返回 needle 在全文文本单元中的出现次数与命中单元下标。"""
    if not needle:
        return 0, []
    hits = []
    for ui, u in enumerate(units):
        start = 0
        while True:
            k = u["text"].find(needle, start)
            if k < 0:
                break
            hits.append((ui, k))
            start = k + 1
    return len(hits), hits


def detect_default(f, text, rec):
    """模板已预置默认值（如保证金 10000）时，单独标记为 DEFAULT，避免误判为空白。"""
    dv = f.get("default")
    if dv is None or text is None:
        return False
    if str(dv) in text:
        rec["status"] = "DEFAULT"
        rec["current"] = str(dv)
        return True
    return False


def locate_para(doc, target):
    """按「记录段落序号 + 锚点」定位段落；序号漂移时回退到锚点第 N 次出现。"""
    anchor = target["anchor"]
    idx = target.get("index", -1)
    occ = target.get("occurrence", 1)
    paras = doc.paragraphs
    if 0 <= idx < len(paras) and anchor in paras[idx].text:
        return idx, paras[idx].text, "index"
    hits = [i for i, p in enumerate(paras) if anchor in p.text]
    if len(hits) >= occ:
        return hits[occ - 1], paras[hits[occ - 1]].text, "anchor"
    if hits:
        return hits[0], paras[hits[0]].text, "anchor-first"
    return None, None, "missing"


def analyse_field(doc, f, units):
    """返回单字段状态字典。"""
    tgt = f["target"]
    rec = {
        "id": f["id"],
        "label": f["label"],
        "where": f.get("where", ""),
        "required": bool(f.get("required", False)),
        "format": f.get("format", ""),
        "example": f.get("example", ""),
        "default": f.get("default"),
    }
    idx, text, how = locate_para(doc, tgt)
    if text is None:
        rec.update(status="MISSING", detail=f"找不到锚点「{tgt['anchor']}」", raw=None)
        return rec

    rec["para_index"] = idx
    rec["locate"] = how
    rec["para_text"] = text

    if f.get("mode") == "regex":
        try:
            m = re.search(f["pattern"], text)
        except re.error as e:
            rec.update(status="SPEC_ERROR", detail=f"正则错误：{e}")
            return rec
        if m:
            raw = m.group(0)
            n, hits = count_occurrences(units, raw)
            rec.update(status="BLANK", raw=raw, match_start=m.start(), match_end=m.end(),
                       doc_occurrences=n, context=text[max(0, m.start() - 12):m.end() + 12])
            detect_default(f, text, rec)
        else:
            rec.update(status="FILLED", raw=None,
                       detail="占位符已不存在（视为已填写）", current=text)
    else:  # append
        anchor = tgt["anchor"]
        m = re.search(re.escape(anchor), text)
        if not m:
            rec.update(status="MISSING", detail=f"锚点不在该段落：{anchor}")
            return rec
        tail = text[m.end():].strip()
        if tail == "":
            rec.update(status="BLANK", raw=anchor, match_start=m.start(), match_end=m.end(),
                       doc_occurrences=1, mode="append", context=text)
        else:
            rec.update(status="FILLED", raw=None, detail="锚点后已有内容（视为已填写）",
                       current=tail)
    return rec


def analyse_cell(doc, c, units):
    rec = {
        "id": c["id"],
        "label": c["label"],
        "where": c.get("where", ""),
        "required": bool(c.get("required", False)),
        "format": c.get("format", ""),
        "example": c.get("example", ""),
        "default": c.get("default"),
        "cell": [c["table"], c["row"], c["col"]],
    }
    try:
        tbl = doc.tables[c["table"]]
        cell = tbl.rows[c["row"]].cells[c["col"]]
    except IndexError:
        rec.update(status="MISSING", detail="表格/行列不存在")
        return rec
    if len(cell.paragraphs) > 1:
        rec.update(status="SPEC_ERROR",
                   detail="该单元格含多个段落；整格改写会吞掉段内换行，需人工确认后再填")
        return rec
    text = cell.text
    rec["para_text"] = text
    if c.get("mode") == "regex":
        m = re.search(c["pattern"], text)
        if m:
            rec.update(status="BLANK", raw=m.group(0), match_start=m.start(),
                       match_end=m.end(), context=text)
        else:
            rec.update(status="FILLED", detail="占位符已不存在（视为已填写）", current=text)
    else:
        anchor = c.get("anchor", "")
        if anchor and anchor in text:
            tail = text.split(anchor, 1)[1].strip()
            if tail == "":
                rec.update(status="BLANK", raw=anchor, match_start=text.find(anchor),
                           match_end=text.find(anchor) + len(anchor), mode="append",
                           context=text)
            else:
                rec.update(status="FILLED", current=tail)
        elif anchor:
            rec.update(status="MISSING", detail=f"单元格内无锚点「{anchor}」")
        else:
            rec.update(status="BLANK", raw="", match_start=0, match_end=0, mode="empty",
                       context="")
    return rec


def analyse_grid(doc, g):
    rec = {
        "id": g["id"],
        "label": g["label"],
        "where": g.get("where", ""),
        "required": bool(g.get("required", False)),
        "unit": g.get("unit", ""),
        "note": g.get("note", ""),
        "table": g["table"],
        "cells": {"quarter": [], "month": []},
    }
    try:
        tbl = doc.tables[g["table"]]
    except IndexError:
        rec.update(status="MISSING", detail="目标表格不存在")
        return rec
    filled = 0
    total = 0
    for key, row in (("quarter", g["rows"]["quarter"]), ("month", g["rows"]["month"])):
        cols = g["quarter_cols"] if key == "quarter" else g["month_cols"]
        for ci in cols:
            try:
                txt = tbl.rows[row].cells[ci].text.strip()
            except IndexError:
                txt = ""
            rec["cells"][key].append(txt)
            total += 1
            if txt:
                filled += 1
    rec["status"] = "BLANK" if filled == 0 else ("PARTIAL" if filled < total else "FILLED")
    rec["filled_cells"] = filled
    rec["total_cells"] = total
    return rec


def scan_unregistered(doc, spec_fields, spec_cells):
    """找出文档中所有未被登记为待填项的空白痕迹。"""
    known = {}
    for f in spec_fields:
        # 段落级已知字段的占位符在填写前可能已被替换，这里只按段落登记，避免误报
        idx = f.get("_resolved_para")
        if idx is not None:
            known.setdefault(("p", idx), []).append(f.get("_raw") or "")
    known_cells = {(c["table"], c["row"], c["col"]) for c in spec_cells}
    found = []
    for pi, p in enumerate(doc.paragraphs):
        text = p.text
        for pat in BLANK_PATTERNS:
            for m in re.finditer(pat, text):
                raw = m.group(0)
                covered = any(raw and raw in (k or "") for k in known.get(("p", pi), []))
                if covered:
                    continue
                found.append({
                    "where": f"正文第 {pi + 1} 段",
                    "para_index": pi,
                    "raw": raw,
                    "context": text[max(0, m.start() - 15):m.end() + 15],
                })
    for ti, tbl in enumerate(doc.tables):
        seen = set()
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                if (ti, ri, ci) in known_cells:
                    continue
                text = cell.text
                for pat in BLANK_PATTERNS:
                    for m in re.finditer(pat, text):
                        found.append({
                            "where": f"表格 {ti + 1} 第 {ri + 1} 行第 {ci + 1} 列",
                            "table": ti, "row": ri, "col": ci,
                            "raw": m.group(0),
                            "context": text[max(0, m.start() - 15):m.end() + 15],
                        })
    return found


def check_locked_content(doc, spec):
    """逐格比对规格里登记的标准条款表（如附件一违约金标准表）。

    scan_blanks.py 与 verify_fill.py 共用：扫描阶段用于模板漂移预警，
    验收阶段用于确认这些条款在成品里一个字都没被改。
    locked_content 为空（公开脱敏版未登记标准表）时本函数为 no-op，
    verify_fill 会显式披露「未登记」，不做静默通过。
    """
    out = []
    for lk in spec.get("locked_content", []):
        ti = lk["table"]
        ok, detail = True, ""
        try:
            tbl = doc.tables[ti]
            got = []
            for row in tbl.rows:
                vals, prev = [], None
                for cell in row.cells:
                    if cell._tc == prev:
                        continue
                    prev = cell._tc
                    vals.append(cell.text)
                got.append(vals)
            exp = lk["expected"]
            if len(got) != len(exp) or any(
                len(g) != len(e) for g, e in zip(got, exp)
            ):
                ok, detail = False, "表结构（行/列）与标准版不一致"
            else:
                for g_row, e_row in zip(got, exp):
                    for gv, ev in zip(g_row, e_row):
                        if gv.strip() != ev.strip():
                            ok, detail = False, (
                                f"期望「{ev[:20]}」，实际「{gv[:20]}」")
                            break
                    if not ok:
                        break
        except IndexError:
            ok, detail = False, "表格不存在"
        out.append({"label": lk["label"], "ok": ok, "detail": detail})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True)
    ap.add_argument("--doc", required=True)
    ap.add_argument("--json-out")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    with open(args.spec, encoding="utf-8") as fh:
        spec = json.load(fh)

    doc = Document(args.doc)
    units = build_units(doc)

    records = []
    for f in spec.get("fields", []):
        r = analyse_field(doc, f, units)
        f["_resolved_para"] = r.get("para_index")
        f["_raw"] = r.get("raw")
        records.append(r)
    for c in spec.get("cells", []):
        records.append(analyse_cell(doc, c, units))
    for g in spec.get("grids", []):
        records.append(analyse_grid(doc, g))

    unregistered = scan_unregistered(doc, spec.get("fields", []), spec.get("cells", []))

    readonly = []
    full_text = "\n".join(u["text"] for u in units)
    for rc in spec.get("readonly_check", []):
        readonly.append({
            "label": rc["label"],
            "ok": rc["text"] in full_text,
            "text": rc["text"],
        })

    locked = check_locked_content(doc, spec)

    summary = {
        "blank": sum(1 for r in records if r["status"] == "BLANK"),
        "partial": sum(1 for r in records if r["status"] == "PARTIAL"),
        "filled": sum(1 for r in records if r["status"] == "FILLED"),
        "default": sum(1 for r in records if r["status"] == "DEFAULT"),
        "missing": sum(1 for r in records if r["status"] == "MISSING"),
        "spec_error": sum(1 for r in records if r["status"] == "SPEC_ERROR"),
        "unregistered": len(unregistered),
        "readonly_fail": sum(1 for r in readonly if not r["ok"]),
        "locked_fail": sum(1 for r in locked if not r["ok"]),
    }

    out = {
        "doc": os.path.abspath(args.doc),
        "spec": spec["doc_key"],
        "title": spec.get("title", ""),
        "scanned_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": summary,
        "records": records,
        "unregistered": unregistered,
        "readonly": readonly,
        "locked": locked,
    }

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump(out, fh, ensure_ascii=False, indent=2)

    if not args.quiet:
        print(f"文档：{os.path.basename(args.doc)}")
        print(f"模板：{spec.get('title')}  ({spec.get('doc_key')})")
        print(f"待填项 {len(records)} 项 | 空白 {summary['blank']} / 部分 {summary['partial']}"
              f" / 默认值 {summary['default']} / 已填 {summary['filled']}"
              f" / 定位失败 {summary['missing']}")
        print("-" * 78)
        for r in records:
            tag = {"BLANK": "空  ", "FILLED": "已填", "MISSING": "异常",
                   "PARTIAL": "部分", "DEFAULT": "默认", "SPEC_ERROR": "配置"}.get(
                       r["status"], r["status"])
            raw = r.get("raw")
            extra = f" 占位符=<{raw}>" if raw else ""
            if r["status"] == "FILLED" and r.get("current"):
                extra = f" 现值=<{r['current'][:24]}>"
            print(f"[{tag}] {r['id']:<7} {r['label']}{extra}")
        if unregistered:
            print("-" * 78)
            print(f"⚠ 未登记空白痕迹 {len(unregistered)} 处：")
            for u in unregistered[:20]:
                print(f"   · {u['where']}  <{u['raw']}>  … {u['context']}")
        if any(not x["ok"] for x in readonly):
            print("-" * 78)
            print("⚠ 甲方固定信息校验未通过：")
            for x in readonly:
                if not x["ok"]:
                    print(f"   · 缺少/被改：{x['label']}  应为「{x['text']}」")
        if any(not x["ok"] for x in locked):
            print("-" * 78)
            print("⚠ 锁定内容与标准版不一致：")
            for x in locked:
                if not x["ok"]:
                    print(f"   · {x['label']}：{x['detail']}")

    # SKILL.md 说「定位失败 / 未登记空白 → 停止」，旧版只打印文字、退出码恒 0，
    # 代理很容易继续往下走。这里把闸门做成机器可判的非零退出码。
    blocked = (summary["missing"] > 0 or summary["spec_error"] > 0
               or summary["unregistered"] > 0 or summary["readonly_fail"] > 0
               or summary["locked_fail"] > 0)
    if blocked and not args.quiet:
        print("-" * 78)
        print("✗ 模板校验未通过：请先修正上述项，不要继续填充。")
    return 1 if blocked else 0


if __name__ == "__main__":
    sys.exit(main())
