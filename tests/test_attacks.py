# -*- coding: utf-8 -*-
"""信任链与越界攻击回归。

这些用例对应第三轮对抗性审计里实证过的绕过路径，任何一个失守都意味着
「用户未确认的内容可能被交付」，因此全部按 fail-closed 断言。

合并仓适配：verify_fill 强制 --review-release（最终三文件放行凭证），
verify() 在每次调用时用 fixtures.make_release_document 生成绑定当前 --result
的真凭证；攻击样本（evil.docx / styles.docx / …）各自绑定自己的凭证，
放行门禁放行后由计划可信性与包级完整性检查完成拦截。
"""
import json
import os
import subprocess
import sys
import zipfile

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
    case = make_case(tmp_path)
    proc = run("build_plan.py", "--spec", case["spec"], "--doc", case["template"],
               "--values", case["values"], "--out", os.path.join(case["dir"], "plan.json"))
    assert proc.returncode == 0, proc.stdout + proc.stderr
    case["plan"] = os.path.join(case["dir"], "plan.json")
    case["result"] = os.path.join(case["dir"], "result.docx")
    fill_sim.fill(case["template"], case["result"], case["plan"])
    return case


def verify(case, result=None, plan=None, spec=None, values=None, extra=()):
    result = result or case["result"]
    release = make_release_document(result, case["spec"], case["values"], case["dir"])
    return run("verify_fill.py", "--template", case["template"],
               "--result", result,
               "--plan", plan or case["plan"],
               "--spec", spec or case["spec"],
               "--values", values or case["values"],
               "--review-release", release, *extra)


def test_baseline_passes(case):
    proc = verify(case)
    assert proc.returncode == 0, proc.stdout + proc.stderr


def test_plan_value_and_expect_tampered_together(case):
    """P0：把 value 与 expect 一起改、并同步改成品 —— 重编译必须发现。"""
    with open(case["plan"], encoding="utf-8") as fh:
        plan = json.load(fh)
    for op in plan["ops"]:
        if op["field_id"] == "M-02":
            op["value"] = op["expect"] = "1元"
    evil_plan = os.path.join(case["dir"], "evil-plan.json")
    with open(evil_plan, "w", encoding="utf-8") as fh:
        json.dump(plan, fh, ensure_ascii=False)

    evil_doc = os.path.join(case["dir"], "evil.docx")
    fill_sim.fill(case["template"], evil_doc, evil_plan)
    proc = verify(case, result=evil_doc, plan=evil_plan)
    assert proc.returncode == 1, proc.stdout
    assert "计划不可信" in proc.stdout


def test_forged_spec_rejected(case):
    """P1：--spec 指向放宽了必填的伪造规格 —— 规格内容指纹必须发现。"""
    with open(case["spec"], encoding="utf-8") as fh:
        spec = json.load(fh)
    spec["fields"][0]["required"] = False
    fake = os.path.join(case["dir"], "fake-spec.json")
    with open(fake, "w", encoding="utf-8") as fh:
        json.dump(spec, fh, ensure_ascii=False)
    proc = verify(case, spec=fake)
    assert proc.returncode == 1, proc.stdout
    assert "规格文件与编译时不一致" in proc.stdout


def test_values_changed_without_recompile(case):
    changed = os.path.join(case["dir"], "values2.json")
    with open(case["values"], encoding="utf-8") as fh:
        values = json.load(fh)
    values["M-02"] = "20000"
    with open(changed, "w", encoding="utf-8") as fh:
        json.dump(values, fh, ensure_ascii=False)
    proc = verify(case, values=changed)
    assert proc.returncode == 1, proc.stdout
    assert "计划不可信" in proc.stdout


def test_spec_and_values_are_required(case):
    no_spec = run("verify_fill.py", "--template", case["template"],
                  "--result", case["result"], "--plan", case["plan"],
                  "--values", case["values"])
    assert no_spec.returncode != 0
    no_values = run("verify_fill.py", "--template", case["template"],
                    "--result", case["result"], "--plan", case["plan"],
                    "--spec", case["spec"])
    assert no_values.returncode != 0


def test_base_option_is_gone(case):
    """--base 曾能把上一轮越界洗白，已移除；传了必须报错。"""
    proc = verify(case, extra=("--base", case["result"]))
    assert proc.returncode != 0
    assert "--base" in proc.stderr


def test_unplanned_text_edit_detected(case):
    with open(case["plan"], encoding="utf-8") as fh:
        plan = json.load(fh)
    evil_plan = os.path.join(case["dir"], "addop.json")
    plan["ops"].append({"field_id": "EVIL", "kind": "insert", "find": "甲方",
                        "occurrence": 1, "value": "（放弃一切追索权）",
                        "expect": "（放弃一切追索权）"})
    with open(evil_plan, "w", encoding="utf-8") as fh:
        json.dump(plan, fh, ensure_ascii=False)
    evil_doc = os.path.join(case["dir"], "evil2.docx")
    fill_sim.fill(case["template"], evil_doc, evil_plan)
    proc = verify(case, result=evil_doc, plan=evil_plan)
    assert proc.returncode == 1, proc.stdout


def test_package_part_tampering_detected(case):
    """包级：往 styles.xml 的 docDefaults 里改默认字体，正文文字一字未动。"""
    evil = os.path.join(case["dir"], "styles.docx")
    with zipfile.ZipFile(case["result"]) as zin, \
            zipfile.ZipFile(evil, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/styles.xml":
                text = data.decode("utf-8")
                original = text
                if '<w:sz w:val="22"/>' in text:
                    text = text.replace('<w:sz w:val="22"/>', '<w:sz w:val="44"/>', 1)
                else:
                    text = text.replace("Calibri", "Times New Roman", 1)
                assert text != original, "styles.xml 未被改动，样本无效"
                data = text.encode("utf-8")
            zout.writestr(item, data)
    proc = verify(case, result=evil)
    assert proc.returncode == 1, proc.stdout
    assert "包级改动" in proc.stdout or "样式表" in proc.stdout


def test_nested_table_text_detected(case):
    """在单元格里塞一个嵌套表格承载额外条款。"""
    evil = os.path.join(case["dir"], "nested.docx")
    with zipfile.ZipFile(case["result"]) as zin, \
            zipfile.ZipFile(evil, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                text = data.decode("utf-8")
                original = text
                nested = ("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>附加条款：乙方无需承担任何责任</w:t>"
                          "</w:r></w:p></w:tc></w:tr></w:tbl>")
                text = text.replace("</w:tc>", nested + "</w:tc>", 1)
                assert text != original, "document.xml 未被改动，样本无效"
                data = text.encode("utf-8")
            zout.writestr(item, data)
    proc = verify(case, result=evil)
    assert proc.returncode == 1, proc.stdout


def test_corrupted_package_fails_cleanly(case):
    """删掉一个被 rels 引用的部件：不得崩溃成 traceback，必须给出明确结论。"""
    evil = os.path.join(case["dir"], "broken.docx")
    with zipfile.ZipFile(case["result"]) as zin, \
            zipfile.ZipFile(evil, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "docProps/core.xml":
                continue
            zout.writestr(item, zin.read(item.filename))
    proc = verify(case, result=evil)
    assert proc.returncode == 1, proc.stdout
    assert "Traceback" not in proc.stderr


def test_apply_requires_protect(case):
    """写入层：--protect 改为必需，省略必须直接报错。"""
    proc = run("apply_fill.py", "--plan", case["plan"], "--edsdk", "/nonexistent/edsdk.py",
               "--file-id", "X", "--spec", case["spec"], "--values", case["values"])
    assert proc.returncode != 0
    assert "--protect" in proc.stderr


def test_template_untouched_by_tests(case):
    """所有测试都只动临时目录；模板内容哈希在测试前后一致。"""
    import hashlib
    before = hashlib.sha256(open(case["template"], "rb").read()).hexdigest()
    verify(case)
    after = hashlib.sha256(open(case["template"], "rb").read()).hexdigest()
    assert before == after
