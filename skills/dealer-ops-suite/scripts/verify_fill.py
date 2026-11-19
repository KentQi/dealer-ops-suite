#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_fill.py —— 填充结果验收（只读，唯一交付闸门）

判定十件事，缺一不可：
  1. 计划可信（重编译）：用「--spec + --values + 原始模板」重新编译一份参考计划，
     逐条核对 plan.ops，并校验 spec 内容指纹、值指纹、source_values。
     这证明「成品 == 计划 == 用户确认值」，而不只是「成品 == 计划」。
  2. 越界（文本）：把计划在**原始模板**上哑演算一次，再与成品逐字比对。
  3. 越界（字符格式）：段落样式 + 字体/字号/加粗/斜体/下划线/颜色集合。
  4. 越界（结构属性）：pPr/tblPr/tblGrid/trPr/tcPr + 分节与纸张。
  5. 越界（页眉页脚文本框）：可见文字 + 域代码序列。
  6. 结构数量：段落数、表格数、文本单元数。
  7. 必填完整性 + 甲方固定信息 + 附件锁定表。
  8. 残留占位符。
  9. 样式表（含 docDefaults）/ 编号定义 / 文档属性：语义指纹一致。
 10. 包级完整性：部件清单一致；除 word/document.xml 外每个部件按命名空间 URI
     规范化后哈希一致；正文里所有结构节点（含单元格内段落 pPr、域代码、
     图形/超链接引用）的子树序列一致。

比对基准**恒为原始模板**。第三轮审计证实：任何让基准脱离原始模板的通道
（旧版 --base）都能把上一轮的越界改动洗白，因此该参数已移除。
二次补填 / 纠错统一用「累积全量 values.json + 原始模板重编译的全量计划」验收。

用法（--spec 与 --values 均为必填）：
  python3 verify_fill.py --template 模板.docx --result 成品.docx \\
      --plan plan.json --spec spec/jx.json --values values.json

退出码：0 通过；1 越界/结构变化/计划不可信；2 仅残留占位符；3 必填未填完。
"""
import argparse
import hashlib
import json
import os
import re
import sys
import zipfile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from scan_blanks import (build_units, iter_blocks,  # noqa: E402
                         analyse_field, analyse_cell, analyse_grid,
                         check_locked_content)
from plan_trust import check_plan, check_plan_selfconsistency  # noqa: E402


# ------------------------------------------------------------------ 索引
def cell_index(units):
    """(table,row,col) → 该单元格第一个文本单元下标"""
    idx = {}
    for i, u in enumerate(units):
        if u["kind"] == "cell":
            idx.setdefault((u["table"], u["row"], u["col"]), i)
    return idx


def find_span(units, needle, occurrence):
    """在有序文本单元里找第 occurrence 次出现，返回 (unit_idx, offset) 或 None"""
    seen = 0
    for i, u in enumerate(units):
        start = 0
        while True:
            k = u["text"].find(needle, start)
            if k < 0:
                break
            seen += 1
            if seen == occurrence:
                return i, k
            start = k + 1
    return None


# ------------------------------------------------------------------ 格式指纹
def _run_fmt(run):
    f = run.font
    size = round(f.size.pt, 1) if f.size is not None else None
    color = None
    try:
        color = str(f.color.rgb) if f.color and f.color.rgb else None
    except (AttributeError, ValueError):
        color = None
    return (f.name, size, bool(f.bold), bool(f.italic),
            f.underline is not None and f.underline is not False, color)


# 这些 run 级属性会改变可见呈现，或把内容藏起来（w:vanish 可承载肉眼看不见的条款），
# 但不在 _run_fmt 的六项里。第三轮审计实证：给已有 run 加 <w:vanish/> 或
# <w:highlight w:val="yellow"/>，文字一个字没变，旧版验收仍判「✓ 通过」。
# 只收「隐藏 / 显著改变可见语义」的属性。纯排版类（w:spacing 字间距、
# w:position 上下标位置、w:w 字符缩放、w:em 等）在整格改写时本来就会被编辑器重建，
# 纳入比对会让合法填充误报；它们由段落/表格结构属性指纹那一层另行覆盖。
RISK_RPR_TAGS = {
    "vanish", "specVanish", "webHidden",     # 隐藏文字：可承载肉眼看不见的额外条款
    "highlight", "shd", "bdr",               # 高亮 / 底纹 / 字符边框
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "effect",
}


def _run_risk(run):
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        return ""
    parts = []
    for c in rPr:
        if not isinstance(c.tag, str):
            continue
        if c.tag.split("}")[-1] in RISK_RPR_TAGS:
            parts.append(_xml(c))
    return "|".join(sorted(parts))


def build_signatures(doc):
    """与 build_units 同序的格式指纹列表：段落样式 + 非空 run 的格式集合"""
    sigs = []
    for kind, blk, _ti in iter_blocks(doc):
        if kind == "p":
            sigs.append(_para_sig(blk))
        else:
            seen = set()
            for row in blk.rows:
                for cell in row.cells:
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    for p in cell.paragraphs:
                        sigs.append(_para_sig(p))
    return sigs


def _para_sig(p):
    style = p.style.name if p.style is not None else ""
    fmts = sorted({_run_fmt(r) for r in p.runs if r.text.strip()},
                  key=lambda x: str(x))
    # 风险属性要对「所有」run 取集合：隐藏文字 run 的文字可能是空的。
    risks = sorted({_run_risk(r) for r in p.runs})
    return f"{style}|{fmts}|{risks}"


# ------------------------------------------------------------------ 结构属性指纹
def _h(s):
    return hashlib.sha256(s.encode("utf-8")).hexdigest()[:16]


# 布尔型 XML 元素：<b/> 与 <b w:val="1"/>（以及 w:val="on"/"true"）语义相同。
BOOL_TAGS = {"b", "bCs", "i", "iCs", "u", "strike", "dstrike", "vanish", "caps",
             "smallCaps", "outline", "shadow", "emboss", "imprint", "noProof",
             "snapToGrid", "widen", "keepNext", "keepLines", "pageBreakBefore",
             "cantSplit", "tblHeader", "hideMark", "adjustRightInd", "widowControl"}


def _xml(e):
    """把元素转成「语义等价」的规范串。

    原模板与 editor_sdk 另存的副本之间存在三类无意义差异，必须归一，否则整篇误报：
      · xmlns 前缀不同          → 只保留本地名
      · 样式写成 name 或 styleId → 跳过 pStyle（样式名由字符指纹那一路比对）
      · <b/> 与 <b w:val="1"/>  → 布尔属性归一
      · rsid* 与 header/footer 的 rId 会被重新编号 → 忽略
    """
    if e is None:
        return ""

    def walk(el):
        if not isinstance(el.tag, str):
            return ""
        tag = el.tag.split("}")[-1]
        if tag == "pStyle":
            return ""
        attrs = []
        for k, v in sorted(el.attrib.items()):
            kk = k.split("}")[-1]
            if kk.startswith("rsid"):
                continue
            if tag in ("headerReference", "footerReference") and kk == "id":
                continue
            if kk == "val" and v in ("true", "1", "on"):
                v = "1"
            if kk == "val" and v in ("false", "0", "off"):
                v = "0"
            attrs.append(f"@{kk}={v}")
        if tag in BOOL_TAGS and not el.attrib:
            attrs.append("@val=1")
        kids = "".join(walk(c) for c in el)
        text = (el.text or "").strip()
        attr_s = "".join(attrs)
        return f"<{tag}{attr_s}{'>' + text if text else ''}>{kids}</{tag}>"

    return walk(e)


def prop_digests(doc):
    """按文档顺序给每个「块」的属性做指纹。

    覆盖文本比对看不见的东西：段落对齐/缩进/间距、段落样式的默认字符格式、
    表格边框、列宽、行高、每格底纹/合并跨度/垂直对齐。
    正文文字改动不会影响这里的任何一项，所以它只会在「格式被动过」时变化。
    """
    out = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            out.append(("段落", _h(_xml(child.find(qn("w:pPr"))))))
        elif child.tag == qn("w:tbl"):
            parts = []
            for tag in ("w:tblPr", "w:tblGrid"):
                parts.append(_xml(child.find(qn(tag))))
            for tr in child.findall(qn("w:tr")):
                parts.append(_xml(tr.find(qn("w:trPr"))))
                for tc in tr.findall(qn("w:tc")):
                    parts.append(_xml(tc.find(qn("w:tcPr"))))
            out.append(("表格", _h("|".join(parts))))
    return out


# ------------------------------------------------------------------ 页眉页脚文本框
def extra_parts(path):
    """抽取页眉/页脚/文本框的「可见文本 + 域代码序列」。

    只比文字会漏掉域被删：删掉 PAGE 域后可见文字仍可能是「第 1 页」，
    但页码不再自动更新。所以 instrText 与 fldChar 的类型序列也要一起比。
    """
    out = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if re.fullmatch(r"word/(header|footer)\d*\.xml", name):
                s = z.read(name).decode("utf-8", "ignore")
                text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", s))
                instr = "|".join(re.findall(r"<w:instrText[^>]*>([^<]*)</w:instrText>", s))
                fld = "|".join(re.findall(r'<w:fldChar[^>]*w:fldCharType="([^"]+)"', s))
                out[name] = f"{text}‖{instr}‖{fld}"
        s = z.read("word/document.xml").decode("utf-8", "ignore")
        tb = "".join(re.findall(r"<w:txbxContent>.*?</w:txbxContent>", s, re.S))
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", tb))
        instr = "|".join(re.findall(r"<w:instrText[^>]*>([^<]*)</w:instrText>", tb))
        fld = "|".join(re.findall(r'<w:fldChar[^>]*w:fldCharType="([^"]+)"', tb))
        out["#textbox"] = f"{text}‖{instr}‖{fld}"
    return out


# ------------------------------------------------------------------ 样式与文档属性
def style_digest(path):
    """样式表 / 编号定义的语义指纹。

    改 styles.xml 里 Normal 的字体，正文 run 一个字都不会变，字符格式指纹也看不到
    （python-docx 的 run.font 只读 run 自己的 rPr）。第三轮实证为盲区，故单独比对。
    _xml 已归一命名空间前缀、属性顺序、rsid 与布尔写法，避免 editor_sdk 另存误报。
    """
    out = []
    with zipfile.ZipFile(path) as z:
        for name in ("word/styles.xml", "word/numbering.xml"):
            if name not in z.namelist():
                continue
            root = etree.fromstring(z.read(name))
            for el in root:
                if not isinstance(el.tag, str):
                    continue
                tag = el.tag.split("}")[-1]
                # docDefaults 定义全篇默认字体/字号，改它比改某个具名样式更隐蔽：
                # 正文 run 里没有任何痕迹，但显示效果全变了。
                # latentStyles 是 Word 自动生成的潜在样式清单，编辑器常改写，不纳入。
                if tag in ("style", "docDefaults", "abstractNum", "num"):
                    out.append(f"{name}:{tag}:{_xml(el)}")
    return _h("|".join(out))


META_VOLATILE = {"created", "modified", "revision", "lastModifiedBy", "lastPrinted"}


def meta_digest(path):
    """文档属性指纹；忽略 Word 每次保存都会重写的时间戳与修订号。"""
    with zipfile.ZipFile(path) as z:
        if "docProps/core.xml" not in z.namelist():
            return ""
        root = etree.fromstring(z.read("docProps/core.xml"))
    parts = []
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        tag = el.tag.split("}")[-1]
        if tag in META_VOLATILE or tag == "coreProperties":
            continue
        parts.append(f"{tag}={(el.text or '').strip()}")
    return _h("|".join(sorted(parts)))


# ------------------------------------------------------------------ 包级完整性
# 第三轮 OOXML 审计的核心结论：正文文本比对之外，DOCX 是一个多部件 ZIP 包。
# 旧版只看 header/footer 的「文字 + 域代码」，styles / settings / numbering /
# rels / Content_Types / docProps / customXml / media / 注释 / 页眉里的表格图片
# 全都不受控。这里改成「部件清单 + 每个部件的语义规范化哈希」，只有
# word/document.xml 因填写而豁免，并另用「结构骨架」盯住它。
APP_VOLATILE = {"Pages", "Words", "Characters", "Lines", "Paragraphs", "TotalTime",
                "ScaleCrop", "LinksUpToDate", "SharedDoc", "HyperlinksChanged",
                "AppVersion", "Template"}

# 合法填充只应改 w:t 文本、必要时拆分/合并 w:r 与 w:rPr；这些节点及其子树不该动。
STRUCT_ROOTS = {"pPr", "tblPr", "tblPrEx", "tblGrid", "trPr", "tcPr", "sectPr",
                "instrText", "fldChar", "fldSimple", "drawing", "pict", "object",
                "hyperlink", "bookmarkStart", "bookmarkEnd", "headerReference",
                "footerReference", "sdt", "docPartObj", "commentRangeStart",
                "commentRangeEnd", "footnoteReference", "endnoteReference"}


def _canon_xml(el, *, sort_kids=False, volatile=(), skip_subtree=(), drop_id=False):
    """语义规范化：按 namespaces URI 比较（换前缀不算改动，换 URI 算改动）。

    忽略 rsid / paraId / textId / Ignorable、布尔属性写法（1/true/on）等
    editor_sdk 另存时的无意义差异。
    """
    if el is None or not isinstance(el.tag, str):
        return ""
    q = etree.QName(el.tag)
    tag, local = q.text, q.localname
    if local in skip_subtree:
        return ""
    attrs = []
    for k, v in sorted(el.attrib.items(), key=lambda x: _ql(x[0])):
        kq = etree.QName(k) if k.startswith("{") else None
        kk = kq.localname if kq else k
        if kk.startswith("rsid") or kk in ("Ignorable", "paraId", "textId"):
            continue
        if drop_id and local == "Relationship" and kk == "Id":
            continue
        if kk == "val" and v in ("true", "1", "on"):
            v = "1"
        if kk == "val" and v in ("false", "0", "off"):
            v = "0"
        attrs.append(f"@{kq.text if kq else k}={v}")
    if local in BOOL_TAGS and not el.attrib:
        # 属性名要与真实 <w:b w:val="1"/> 的带命名空间写法一致，否则
        # 「<w:b/>」与「<w:b w:val="on"/>」会被误判为不同。
        ns = q.namespace or ""
        key = "{" + ns + "}val" if ns else "val"
        attrs.append(f"@{key}=1")
    text = "" if local in volatile else (el.text or "").strip()
    kids = [_canon_xml(c, sort_kids=sort_kids, volatile=volatile,
                       skip_subtree=skip_subtree, drop_id=drop_id) for c in el]
    if sort_kids:
        kids.sort()
    return f"<{tag}{''.join(attrs)}{'>' + text if text else ''}>{''.join(kids)}</{tag}>"


def _ql(k):
    return etree.QName(k).localname if k.startswith("{") else k


def _doc_struct(data):
    """正文结构骨架：关键结构节点的规范化子树 + 表格嵌套形态。

    表格形态单独记：嵌套表格里的文字不会被 build_units 抓到（它只遍历 body
    直接子元素与 Table.rows），第三轮审计实证「在单元格里塞一个含附加条款的
    嵌套表格」可以完全隐身。记下每张表的嵌套深度即可发现新增/嵌套。
    """
    root = etree.fromstring(data)
    out = []
    shape = []

    def walk(el, depth=0):
        if not isinstance(el.tag, str):
            return
        local = etree.QName(el.tag).localname
        if local == "tbl":
            shape.append(depth)
            depth += 1
        if local in STRUCT_ROOTS:
            out.append(_canon_xml(el))
            return
        for c in el:
            walk(c, depth)

    walk(root)
    return _h("|".join(out)) + ":" + ",".join(str(x) for x in shape)


def package_fingerprint(path):
    """返回 (部件名 -> 规范化指纹, 正文结构骨架指纹)。"""
    parts = {}
    struct = ""
    with zipfile.ZipFile(path) as z:
        for n in sorted(z.namelist()):
            if n.endswith("/"):
                continue
            data = z.read(n)
            if n == "word/document.xml":
                struct = _doc_struct(data)
                parts[n] = "(正文另由文本/格式/结构骨架比对)"
                continue
            try:
                if n == "[Content_Types].xml":
                    s = _canon_xml(etree.fromstring(data), sort_kids=True)
                elif n == "docProps/core.xml":
                    s = _canon_xml(etree.fromstring(data), volatile=META_VOLATILE)
                elif n == "docProps/app.xml":
                    s = _canon_xml(etree.fromstring(data), volatile=APP_VOLATILE)
                elif n.endswith(".rels"):
                    s = _canon_xml(etree.fromstring(data), sort_kids=True, drop_id=True)
                elif n == "word/settings.xml":
                    s = _canon_xml(etree.fromstring(data), skip_subtree=("rsids",))
                elif n.endswith(".xml"):
                    s = _canon_xml(etree.fromstring(data))
                else:
                    s = hashlib.sha256(data).hexdigest()
            except (etree.XMLSyntaxError, ValueError):
                s = hashlib.sha256(data).hexdigest()
            parts[n] = _h(s)
    return parts, struct


# ------------------------------------------------------------------ 哑演算
def simulate(base_units, ops):
    """在基准文本上哑演算计划：先在原文解析全部位置，再从后往前套用。"""
    texts = [u["text"] for u in base_units]
    idx_map = cell_index(base_units)
    errors = []
    edits = {}

    for op in ops:
        fid = op.get("field_id", "?")
        if op["kind"] == "cell":
            key = (op["table"], op["row"], op["col"])
            i = idx_map.get(key)
            if i is None:
                errors.append(f"{fid}：基准中无单元格 表#{op['table'] + 1} "
                              f"R{op['row'] + 1}C{op['col'] + 1}")
                continue
            edits.setdefault(i, []).append((0, len(texts[i]), op["value"]))
        elif op["kind"] in ("insert", "replace"):
            hit = find_span(base_units, op["find"], int(op.get("occurrence", 1)))
            if not hit:
                errors.append(f"{fid}：基准中找不到「{op['find']}」")
                continue
            i, k = hit
            if op["kind"] == "insert":
                cut = k + len(op["find"])
                edits.setdefault(i, []).append((cut, cut, op["value"]))
            else:
                edits.setdefault(i, []).append((k, k + len(op["find"]), op["value"]))
        else:
            errors.append(f"{fid}：未知操作类型 {op['kind']}")

    for i, spans in edits.items():
        for start, end, value in sorted(spans, key=lambda s: s[0], reverse=True):
            texts[i] = texts[i][:start] + value + texts[i][end:]
    return texts, errors


def check_plan_trust(plan, spec_path, values_path, template):
    """计划可信性 = 内部自洽 + 「由这份 spec / values / 原始模板重编译」复核。

    只做内部自洽是第三轮审计发现的 P0：value 与 expect 都在同一份 plan.json 里，
    一起改就能让验收把用户从未确认的条款放行。这里必须再重编译一次。
    """
    problems = check_plan_selfconsistency(plan)
    problems += check_plan(plan, spec_path, values_path, template)
    return problems


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--result", required=True)
    ap.add_argument("--plan", required=True)
    ap.add_argument("--spec", required=True, help="字段规格 JSON（验收的可信根之一）")
    ap.add_argument("--values", required=True, help="用户确认值 JSON（验收的可信根之一）")
    args = ap.parse_args()

    if os.path.realpath(args.template) == os.path.realpath(args.result):
        print("✗ 模板与成品是同一个文件，比对无意义。请检查参数。")
        return 1
    for p, label in ((args.spec, "--spec"), (args.values, "--values")):
        if not os.path.exists(p):
            print(f"✗ {label} 指向的文件不存在：{p}")
            return 1

    try:
        tdoc = Document(args.template)
        rdoc = Document(args.result)
    except Exception as e:
        # 删掉 footer part、把 rels 指向不存在的部件等，会让 python-docx 直接抛
        # KeyError；旧版因此崩溃而非给出「未通过」，第三轮审计列为可用性缺陷。
        print(f"✗ 无法读取文档（DOCX 包或关系引用可能已损坏）：{type(e).__name__}: {e}")
        print("  比对无法完成，禁止交付。请用模板备份还原后重来。")
        return 1
    tu = build_units(tdoc)
    ru = build_units(rdoc)
    plan = json.load(open(args.plan, encoding="utf-8"))
    spec = json.load(open(args.spec, encoding="utf-8"))
    ops = plan.get("ops", [])

    ok = True
    struct = []

    # 1) 结构（必须计入 hard：旧版只置 ok=False，而 ok 从未参与最终裁定）
    if len(tu) != len(ru):
        msg = f"结构变化：模板 {len(tu)} 个文本单元，成品 {len(ru)} 个"
        print("✗ " + msg)
        struct.append(msg)
        ok = False
    if len(tdoc.tables) != len(rdoc.tables):
        msg = f"表格数量变化：模板 {len(tdoc.tables)} → 成品 {len(rdoc.tables)}"
        print("✗ " + msg)
        struct.append(msg)
        ok = False
    if len(tdoc.paragraphs) != len(rdoc.paragraphs):
        msg = f"段落数量变化：模板 {len(tdoc.paragraphs)} → 成品 {len(rdoc.paragraphs)}"
        print("✗ " + msg)
        struct.append(msg)
        ok = False

    # 2) 计划可信性（重编译，锚定用户确认值）
    trust = check_plan_trust(plan, args.spec, args.values, args.template)
    if plan.get("doc") != spec.get("doc_key"):
        trust.append(
            f"计划标注的合同类型「{plan.get('doc')}」与规格「{spec.get('doc_key')}」不一致，"
            "疑似拿错合同 / 拿错规格")
    if trust:
        ok = False
        print("✗ 计划不可信：")
        for t in trust[:10]:
            print("   ·", t)

    # 3) 文本哑演算：比对基准恒为原始模板。
    #    旧版 --base 让基准脱离模板后，上一轮的越界文本会被当成「已确认填入」洗白，
    #    第三轮审计已实证（--base=成品 + 空计划 → 退出码 0），故该通道已移除。
    bu = tu
    sim, sim_errors = simulate(bu, ops)
    if sim_errors:
        print("✗ 计划本身有问题，无法完成比对：")
        for e in sim_errors:
            print("   ·", e)
        ok = False

    diffs = []
    sim_short = False
    if not sim_errors and len(sim) != len(ru):
        # 旧版这里静默跳过比对，若结构检查又没计入 hard，就会出现「比对没做却判通过」。
        sim_short = True
        print(f"✗ 哑演算结果与成品的文本单元数不一致（{len(sim)} vs {len(ru)}），比对未完成")
    elif not sim_errors:
        for i, expect in enumerate(sim):
            actual = ru[i]["text"]
            if expect != actual:
                diffs.append((i, bu[i], expect, actual))
    if diffs:
        ok = False
        print(f"✗ 发现 {len(diffs)} 处计划外改动（文本）：")
        for i, unit, expect, actual in diffs[:20]:
            where = (f"表#{unit['table'] + 1} R{unit['row'] + 1}C{unit['col'] + 1}"
                     if unit["kind"] == "cell" else f"第 {unit['para_index'] + 1} 段")
            print(f"   · {where}")
            print(f"     应为：{expect[:70]}")
            print(f"     实际：{actual[:70]}")
        if len(diffs) > 20:
            print(f"   … 另有 {len(diffs) - 20} 处")

    # 4) 格式指纹
    tsig = build_signatures(tdoc)
    rsig = build_signatures(rdoc)
    fdiffs = []
    if len(tsig) == len(rsig):
        for i, (a, b) in enumerate(zip(tsig, rsig)):
            if a == b:
                continue
            # 原本空白的格子/段落，填入内容后必然多出 run，格式集合变化属预期，不算越界；
            # 只对「原本有字」的位置追究格式是否被动过。
            if tu[i]["text"].strip() == "":
                continue
            unit = tu[i]
            where = (f"表#{unit['table'] + 1} R{unit['row'] + 1}C{unit['col'] + 1}"
                     if unit["kind"] == "cell" else f"第 {unit['para_index'] + 1} 段")
            fdiffs.append((where, a, b))
    else:
        fdiffs.append(("全文", f"签名条目 {len(tsig)}", f"签名条目 {len(rsig)}"))
    if fdiffs:
        ok = False
        print(f"✗ 发现 {len(fdiffs)} 处格式改动（文本可能没变，但样式被动过）：")
        for where, a, b in fdiffs[:15]:
            print(f"   · {where}")
            print(f"     原：{a[:80]}")
            print(f"     现：{b[:80]}")
        if len(fdiffs) > 15:
            print(f"   … 另有 {len(fdiffs) - 15} 处")

    # 5) 页眉页脚文本框（含域代码）
    try:
        tp, rp = extra_parts(args.template), extra_parts(args.result)
    except (KeyError, OSError, zipfile.BadZipFile, ValueError) as e:
        print(f"✗ 页眉页脚部件读取失败，比对无法完成：{type(e).__name__}: {e}")
        return 1
    xdiffs = [(k, tp.get(k, ""), rp.get(k, "")) for k in set(tp) | set(rp)
              if tp.get(k, "") != rp.get(k, "")]
    if xdiffs:
        ok = False
        print(f"✗ 页眉/页脚/文本框被改动 {len(xdiffs)} 处：")
        for k, a, b in xdiffs:
            print(f"   · {k}\n     原：{a[:60]}\n     现：{b[:60]}")

    # 5b) 结构属性（边框 / 列宽 / 行高 / 底纹 / 对齐 / 缩进 / 段落默认字体）
    tprop, rprop = prop_digests(tdoc), prop_digests(rdoc)
    pdiffs = []
    if len(tprop) == len(rprop):
        for i, ((ka, va), (_kb, vb)) in enumerate(zip(tprop, rprop)):
            if va != vb:
                pdiffs.append((i + 1, ka))
    else:
        pdiffs.append((0, f"块数 {len(tprop)} → {len(rprop)}"))
    if pdiffs:
        ok = False
        print(f"✗ 发现 {len(pdiffs)} 处格式/结构改动（边框、列宽、行高、底纹、对齐、缩进等）：")
        for idx, kind in pdiffs[:15]:
            print(f"   · 文档第 {idx} 个块（{kind}）")
        if len(pdiffs) > 15:
            print(f"   … 另有 {len(pdiffs) - 15} 处")

    # 5c) 分节与纸张（页边距、纸张大小、分节符）
    tsect = [_h(_xml(s._sectPr)) for s in tdoc.sections]
    rsect = [_h(_xml(s._sectPr)) for s in rdoc.sections]
    scale_changed = tsect != rsect
    if scale_changed:
        ok = False
        print(f"✗ 分节/纸张设置被改动（{len(tsect)} 节 → {len(rsect)} 节）")

    # 5d) 样式表 / 编号定义 / 文档属性（正文文本与 run 格式都看不见这一层）
    asset_changed = (style_digest(args.template) != style_digest(args.result)
                     or meta_digest(args.template) != meta_digest(args.result))
    if asset_changed:
        ok = False
        print("✗ 样式表 / 编号定义 / 文档属性被改动")

    # 5e) 包级完整性：部件清单 + 非正文部件规范化哈希 + 正文结构骨架
    try:
        tparts, tstruct = package_fingerprint(args.template)
        rparts, rstruct = package_fingerprint(args.result)
    except (KeyError, OSError, zipfile.BadZipFile, etree.XMLSyntaxError, ValueError) as e:
        print(f"✗ 包级完整性检查无法完成：{type(e).__name__}: {e}")
        print("  比对未完成，禁止交付。请用模板备份还原后重来。")
        return 1
    pkg_bad = []
    for n in sorted(set(rparts) - set(tparts)):
        pkg_bad.append(f"新增部件 {n}")
    for n in sorted(set(tparts) - set(rparts)):
        pkg_bad.append(f"缺失部件 {n}")
    for n in sorted(set(tparts) & set(rparts)):
        if n != "word/document.xml" and tparts[n] != rparts[n]:
            pkg_bad.append(f"部件被改 {n}")
    if tstruct != rstruct:
        pkg_bad.append("正文结构骨架被改（单元格段落属性 / 域代码 / 图形与超链接引用等）")
    if pkg_bad:
        ok = False
        print(f"✗ 发现 {len(pkg_bad)} 处包级改动：")
        for p in pkg_bad[:12]:
            print("   ·", p)
        if len(pkg_bad) > 12:
            print(f"   … 另有 {len(pkg_bad) - 12} 处")

    # 6) 残留占位符
    residual = []
    from scan_blanks import BLANK_PATTERNS
    # 排除 X/x 占位符：真实信用代码或型号里可能含 XXX，作为"残留"判据误报率太高
    pats = [p for p in BLANK_PATTERNS if not p.lstrip("(?<![A-Za-z0-9])").startswith(("X", "x"))]
    pats.append(r"(?<!＿)＿(?!＿)")
    for u in ru:
        for p in pats:
            for m in re.finditer(p, u["text"]):
                where = (f"表#{u['table'] + 1} R{u['row'] + 1}C{u['col'] + 1}"
                         if u["kind"] == "cell" else f"第 {u['para_index'] + 1} 段")
                residual.append((where, m.group(0),
                                 u["text"][max(0, m.start() - 18):m.end() + 18]))

    # 7) 甲方固定信息 + 必填完整性 + 附件锁定表
    full = "\n".join(u["text"] for u in ru)
    fixed_missing, missing_required, locked_bad = [], [], []
    for rc in spec.get("readonly_check", []):
        if rc["text"] not in full:
            fixed_missing.append(rc)
    # 必填项不能只靠「占位符残留」发现：追加类字段（联系人：、收货地址：）模板里
    # 本就没有下划线。旧版只认 BLANK，遇到 MISSING / SPEC_ERROR / PARTIAL 会漏判。
    # 只有 FILLED 与 DEFAULT（模板预置值）算完成。
    DONE = ("FILLED", "DEFAULT")
    for f in spec.get("fields", []):
        if f.get("required") and analyse_field(rdoc, f, ru)["status"] not in DONE:
            missing_required.append(f"{f['id']} {f['label']}（{f.get('where', '')}）")
    for c in spec.get("cells", []):
        if c.get("required") and analyse_cell(rdoc, c, ru)["status"] not in DONE:
            missing_required.append(f"{c['id']} {c['label']}")
    for g in spec.get("grids", []):
        if g.get("required"):
            rec = analyse_grid(rdoc, g)
            # 目标表按设计「季度 / 月度二选一档」，PARTIAL 属正常完成状态。
            if rec.get("status") == "MISSING" or rec.get("filled_cells", 0) == 0:
                missing_required.append(f"{g['id']} {g['label']}（整表空白）")
    locked_bad = [x for x in check_locked_content(rdoc, spec) if not x["ok"]]
    if missing_required or fixed_missing or locked_bad:
        ok = False

    print("=" * 78)
    print(f"基准：{os.path.basename(args.template)}（恒为原始模板）")
    print(f"成品：{os.path.basename(args.result)}")
    print(f"计划操作：{len(ops)} 条")
    print(f"结构：{'一致' if len(tu) == len(ru) else '变化'}（文本单元 {len(tu)} → {len(ru)}）")
    print(f"计划可信：{'是' if not trust else '否'}")
    print(f"文本越界：{len(diffs)} 处")
    print(f"字符格式越界：{len(fdiffs)} 处")
    print(f"结构属性越界：{len(pdiffs)} 处")
    print(f"页眉页脚文本框：{'未动' if not xdiffs else str(len(xdiffs)) + ' 处改动'}")
    print(f"样式表与文档属性：{'未动' if not asset_changed else '被改动'}")
    print(f"包级完整性：{'未动' if not pkg_bad else str(len(pkg_bad)) + ' 处改动'}")
    print(f"残留占位符：{len(residual)} 处")
    print(f"必填项完成：{'是' if not missing_required else '否（缺 ' + str(len(missing_required)) + ' 项）'}")
    print(f"甲方固定信息：{'完整' if not fixed_missing else '缺失 ' + str(len(fixed_missing)) + ' 项'}")
    print(f"锁定条款表：{'完整' if not locked_bad else '不一致 ' + str(len(locked_bad)) + ' 项'}")
    print(f"计划重编译：{'通过（与确认值逐条一致）' if not trust else '未通过'}")

    if missing_required:
        print("-" * 78)
        print("必填项仍未填写：")
        for m in missing_required[:20]:
            print(f"   · {m}")
        if len(missing_required) > 20:
            print(f"   … 另有 {len(missing_required) - 20} 项")

    if residual:
        print("-" * 78)
        for w, raw, ctx in residual[:20]:
            print(f"   · 待补 {w}  <{raw}>  … {ctx}")
    if fixed_missing:
        print("-" * 78)
        for rc in fixed_missing:
            print(f"   · 甲方信息缺失：{rc['label']}（应为「{rc['text']}」）")
    if locked_bad:
        print("-" * 78)
        for x in locked_bad:
            print(f"   · 锁定条款表被改：{x['label']}（{x['detail']}）")

    print("=" * 78)
    hard = bool(struct or diffs or fdiffs or pdiffs or xdiffs or trust
                or fixed_missing or locked_bad or sim_errors or sim_short
                or scale_changed or asset_changed or pkg_bad)
    if hard:
        why = []
        if struct:
            why.append("结构数量变化")
        if sim_errors:
            why.append("计划无法在模板上定位（比对未完成）")
        if sim_short:
            why.append("哑演算与成品规模不符（比对未完成）")
        if diffs:
            why.append(f"文本越界 {len(diffs)} 处")
        if fdiffs:
            why.append(f"字符格式越界 {len(fdiffs)} 处")
        if pdiffs:
            why.append(f"结构属性越界 {len(pdiffs)} 处")
        if xdiffs:
            why.append(f"页眉页脚文本框改动 {len(xdiffs)} 处")
        if trust:
            why.append("计划不可信")
        if fixed_missing:
            why.append("甲方信息缺失")
        if locked_bad:
            why.append("锁定条款表被改")
        if scale_changed:
            why.append("分节/纸张设置被改")
        if asset_changed:
            why.append("样式表/编号定义/文档属性被改")
        if pkg_bad:
            why.append(f"包级改动 {len(pkg_bad)} 处")
        print("结论：✗ 未通过 —— " + "、".join(why) + "。请用模板备份还原后重来。")
        return 1
    if missing_required:
        print(f"结论：✗ 未填完 —— 结构无越界，但 {len(missing_required)} 项必填仍空白。"
              "补齐后重跑验收，不要还原。")
        return 3
    if residual:
        print("结论：△ 待补 —— 除计划项外无越界改动，但仍有残留占位符。")
        return 2
    print("结论：✓ 通过 —— 文本、格式、页眉页脚均未越界，必填项齐全，除已确认的填写项外与原模板一致。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
