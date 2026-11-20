# -*- coding: utf-8 -*-
"""按 plan.ops 在 docx 上模拟 editor_sdk 的写入，仅用于测试。

生产环境必须走 editor_sdk（SKILL.md 的硬边界），这里只是为了在 CI 里
不依赖宿主机的 Office 组件就能验证「计划 → 成品 → 验收」这条链。
"""
import json

from docx import Document


def _all_paragraphs(doc):
    out = list(doc.paragraphs)
    for table in doc.tables:
        seen = set()
        for row in table.rows:
            for cell in row.cells:
                if cell._tc in seen:
                    continue
                seen.add(cell._tc)
                out.extend(cell.paragraphs)
    return out


def _set_text(paragraph, text):
    """把整段文本写进第一个 run，其余 run 清空（保留首个 run 的格式）。"""
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(text)
        return
    runs[0].text = text
    for run in runs[1:]:
        run.text = ""


def _find(paragraphs, needle, occurrence):
    seen = 0
    for i, p in enumerate(paragraphs):
        start = 0
        while True:
            k = p.text.find(needle, start)
            if k < 0:
                break
            seen += 1
            if seen == occurrence:
                return i, k
            start = k + 1
    raise AssertionError(f"填充模拟找不到「{needle}」第 {occurrence} 处")


def fill(template, out, plan_path):
    doc = Document(template)
    with open(plan_path, encoding="utf-8") as fh:
        plan = json.load(fh)
    paras = _all_paragraphs(doc)
    for op in plan.get("ops", []):
        if op["kind"] == "cell":
            cell = doc.tables[op["table"]].rows[op["row"]].cells[op["col"]]
            _set_text(cell.paragraphs[0], op["value"])
            continue
        idx, off = _find(paras, op["find"], int(op.get("occurrence", 1)))
        p = paras[idx]
        text = p.text
        if op["kind"] == "insert":
            cut = off + len(op["find"])
            _set_text(p, text[:cut] + op["value"] + text[cut:])
        elif op["kind"] == "replace":
            _set_text(p, text[:off] + op["value"] + text[off + len(op["find"]):])
        else:
            raise AssertionError(f"未知操作类型 {op['kind']}")
    doc.save(out)
    return out
