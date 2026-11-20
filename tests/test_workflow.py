# -*- coding: utf-8 -*-
"""端到端：编译计划 → 模拟填充 → 收尾验收。

合并仓适配：verify_fill 强制 --review-release，verify() 用
fixtures.make_release_document 生成绑定本次 --result 的真凭证。
"""
import json
import os
import subprocess
import sys

import pytest

from fixtures import SCRIPTS, make_case, make_release_document
import fill_sim


def run(script, *args):
    return subprocess.run(
        [sys.executable, str(SCRIPTS / script), *args],
        capture_output=True, text=True,
    )


@pytest.fixture()
def case(tmp_path):
    return make_case(tmp_path)


def build_plan(case, out=None, values=None):
    out = out or os.path.join(case["dir"], "plan.json")
    return run("build_plan.py", "--spec", case["spec"], "--doc", case["template"],
               "--values", values or case["values"], "--out", out), out


def verify(case, result, plan, spec=None, values=None):
    release = make_release_document(result, spec or case["spec"],
                                    values or case["values"], case["dir"])
    return run("verify_fill.py", "--template", case["template"], "--result", result,
               "--plan", plan, "--spec", spec or case["spec"],
               "--values", values or case["values"],
               "--review-release", release)


def test_build_plan_ok(case):
    proc, out = build_plan(case)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    with open(out, encoding="utf-8") as fh:
        plan = json.load(fh)
    assert len(plan["ops"]) == 3
    assert {"spec_sha256", "base_sha256", "values_digest", "source_values"} <= set(plan)
    for op in plan["ops"]:
        assert op["expect"] == op["value"]


def test_end_to_end_verify_passes(case):
    proc, plan = build_plan(case)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    result = os.path.join(case["dir"], "result.docx")
    fill_sim.fill(case["template"], result, plan)
    proc = verify(case, result, plan)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "✓ 通过" in proc.stdout
    assert "计划重编译：通过" in proc.stdout
    assert "包级完整性：未动" in proc.stdout


def test_written_values_really_land(case):
    proc, plan = build_plan(case)
    assert proc.returncode == 0
    result = os.path.join(case["dir"], "result.docx")
    fill_sim.fill(case["template"], result, plan)
    from docx import Document
    doc = Document(result)
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "合同编号：MINI-2026-001" in text
    assert "保证金：10000元" in text
    assert "联系人：张三" in doc.tables[0].rows[0].cells[1].text


def test_unknown_field_id_fails_closed(case):
    bad = os.path.join(case["dir"], "bad-values.json")
    with open(bad, "w", encoding="utf-8") as fh:
        json.dump({"M-01": "MINI-2026-001", "M-99": "x"}, fh, ensure_ascii=False)
    proc, _ = build_plan(case, values=bad)
    assert proc.returncode != 0
    assert "无法识别的字段 ID" in proc.stdout + proc.stderr


def test_missing_required_fails_closed(case):
    bad = os.path.join(case["dir"], "missing.json")
    with open(bad, "w", encoding="utf-8") as fh:
        json.dump({"M-01": "MINI-2026-001"}, fh, ensure_ascii=False)
    proc, _ = build_plan(case, values=bad)
    assert proc.returncode == 1
    assert "必填项未提供" in proc.stdout
