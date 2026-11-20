# -*- coding: utf-8 -*-
"""测试夹具：合成一份迷你合同，不依赖任何真实客户合同或模板。

迷你合同刻意包含三种落笔形态，覆盖 build_plan / apply_fill / verify_fill 的三条代码路径：
  · 段落追加（append）：模板里本来没有下划线，漏填无残留 → 靠必填完整性兜底
  · 段落正则替换（regex）：填占位符
  · 单元格追加（cell）：按表格行列定位

合并仓适配（相对原仓的两处）：
  · SCRIPTS 指向 skills/dealer-ops-suite/scripts（原仓假设 scripts 与 tests 平级）；
  · 新增 make_release_document：verify_fill 现已强制 --review-release（最终三文件
    放行凭证），测试统一用它生成**真凭证**（contract_release.build_release），
    而不是伪造 JSON——伪造会在 verify_final_release 的哈希与路径绑定上被拒。
"""
import json
import os
import shutil
import sys
from pathlib import Path

from docx import Document

# 路径注入写法与 tests/test_contract_fill_gates.py 保持一致。
ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"
sys.path.insert(0, str(SCRIPTS))

import contract_release  # noqa: E402  依赖上一行的 sys.path 注入

# 人读文档目录（skills/dealer-ops-suite/references/），spec 一致性测试用。
REFERENCES = SCRIPTS.parent / "references"

MINI_SPEC = {
    "doc_key": "MINI",
    "title": "测试用迷你合同",
    "fields": [
        {
            "id": "M-01",
            "label": "合同编号",
            "where": "第 1 段「合同编号：」",
            "required": True,
            "format": "1-32 字",
            "example": "MINI-2026-001",
            "target": {"kind": "para", "index": 0, "anchor": "合同编号：", "occurrence": 1},
            "mode": "append",
            "render": "{v}",
            "vtype": "text",
        },
        {
            "id": "M-02",
            "label": "履约保证金",
            "where": "第 2 段「保证金：＿＿＿元」",
            "required": True,
            "format": "1-9 位整数",
            "example": "10000",
            "target": {"kind": "para", "index": 1, "anchor": "保证金：", "occurrence": 1},
            "mode": "regex",
            "pattern": "＿{3,}元",
            "render": "{v}元",
            "vtype": "amount_num",
        },
    ],
    "cells": [
        {
            "id": "M-C1",
            "label": "联系人",
            "where": "表格 1 第 1 行第 2 列「联系人：」",
            "required": True,
            "format": "姓名 2-15 字",
            "example": "张三",
            "table": 0,
            "row": 0,
            "col": 1,
            "mode": "append",
            "anchor": "联系人：",
            "render": "{v}",
            "vtype": "person",
        }
    ],
    "grids": [],
    "readonly_check": [{"label": "甲方名称", "text": "测试甲方有限公司"}],
    "locked_content": [],
}

MINI_VALUES = {"M-01": "MINI-2026-001", "M-02": "10000", "M-C1": "张三"}


def build_mini_template(path):
    """生成迷你合同模板。"""
    doc = Document()
    doc.add_paragraph("合同编号：")
    doc.add_paragraph("保证金：＿＿＿元")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "乙方信息"
    table.cell(0, 1).text = "联系人："
    table.cell(1, 0).text = "甲方"
    table.cell(1, 1).text = "测试甲方有限公司"
    doc.save(path)
    return path


def make_case(tmp_path):
    """在临时目录里铺好 template / spec / values，返回路径字典。"""
    tmp_path = str(tmp_path)
    tpl = os.path.join(tmp_path, "mini.docx")
    spec = os.path.join(tmp_path, "mini-spec.json")
    values = os.path.join(tmp_path, "values.json")
    build_mini_template(tpl)
    with open(spec, "w", encoding="utf-8") as fh:
        json.dump(MINI_SPEC, fh, ensure_ascii=False, indent=2)
    with open(values, "w", encoding="utf-8") as fh:
        json.dump(MINI_VALUES, fh, ensure_ascii=False, indent=2)
    return {"dir": tmp_path, "template": tpl, "spec": spec, "values": values}


def make_release_document(result, spec, values, workdir=None):
    """生成一份**真实的**最终三文件放行凭证并落盘，返回凭证 JSON 路径。

    verify_fill.py 已强制 --review-release，且内部走
    contract_release.verify_final_release：凭证须绑定
      · 三份互不相同的输出文件（main / anti_diversion / price_policy），
        main.path 的 realpath 必须等于本次 verify_fill 的 --result；
      · 五类审阅输入（spec/values/fc_spec/fc_values/policy）的路径 + sha256。
    因此：
      · 附件 / 价格政策用成品的两个**独立副本**占位（内容相同、路径不同即可，
        build_release 只要求三份文件路径互不相同）；
      · 输入绑定当前 case 的 spec / values —— 这两份文件在凭证生成后不得再改
        （本套测试对它们只读；伪造 spec / 改值用例都写成**新文件**）；
      · 每次验证调用都重新生成凭证，保证 main 绑定 = 本次 --result。
    """
    workdir = workdir or os.path.dirname(os.path.abspath(result))
    main_file = os.path.abspath(result)
    fc_file = os.path.join(workdir, "release-anti-diversion-copy.docx")
    policy_file = os.path.join(workdir, "release-price-policy-copy.docx")
    shutil.copyfile(main_file, fc_file)
    shutil.copyfile(main_file, policy_file)
    release = contract_release.build_release(
        [], {"status": "merged"}, main_file, fc_file, policy_file, None,
        inputs={"spec": spec, "values": values,
                "fc_spec": spec, "fc_values": values, "policy": values})
    document = {"schema": contract_release.REVIEW_SCHEMA, "findings": [],
                "llm_review": {"status": "merged"}, "release": release}
    path = os.path.join(workdir, "release.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(document, fh, ensure_ascii=False, indent=2)
    return path
