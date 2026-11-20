#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合同写前预检与最终三文件放行的双门禁验收。"""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"
sys.path.insert(0, str(SCRIPTS))
import contract_release


class ContractFillGateTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(prefix=".test-fill-gates-", dir=ROOT)
        self.root = Path(self._tmp.name)
        self.main = self.root / "主合同.docx"
        self.fc = self.root / "防窜货附件.docx"
        self.policy = self.root / "价格政策.xlsx"
        self.main.write_text("主合同最终版", encoding="utf-8")
        self.fc.write_text("防窜货最终版", encoding="utf-8")
        wb = openpyxl.Workbook()
        for row in (["商品条码", "商品名称", "结算价", "生效起", "生效止"],
                    ["690000000001", "商品", 50, "2026-10-01", None]):
            wb.active.append(row)
        wb.save(self.policy)
        self.values = self.root / "values.json"
        self.values.write_text("{}", encoding="utf-8")
        self.fc_values = self.root / "fc-values.json"
        self.fc_values.write_text('{"FC-C1":"2026-10-01","FC-C2":"2026-10-01"}', encoding="utf-8")

    def tearDown(self):
        self._tmp.cleanup()

    def review_command(self, out, llm, approve=True, resolutions=None):
        args = [sys.executable, str(SCRIPTS / "review_contract_package.py"),
                "--spec", str(self.values), "--values", str(self.values),
                "--fc-spec", str(self.values), "--fc-values", str(self.fc_values),
                "--policy", str(self.policy), "--llm-findings", str(llm), "--out", str(out)]
        if approve:
            args.append("--approve-write")
        else:
            args += ["--release", "--main-file", str(self.main), "--fc-file", str(self.fc)]
        if resolutions:
            args += ["--resolutions", str(resolutions)]
        return subprocess.run(args, text=True, capture_output=True, timeout=20)

    def release_inputs(self, policy=None):
        return {"spec": str(self.values), "values": str(self.values),
                "fc_spec": str(self.values), "fc_values": str(self.fc_values),
                "policy": str(policy or self.policy)}

    def final_document(self, main=None, fc=None, policy=None):
        main, fc, policy = main or self.main, fc or self.fc, policy or self.policy
        findings = []
        release = contract_release.build_release(findings, {"status": "merged"}, str(main), str(fc),
                                                 str(policy), None, inputs=self.release_inputs(policy))
        return {"schema": "dealer-ops/contract-review@1", "findings": findings,
                "llm_review": {"status": "merged"}, "release": release}

    def test_release_requires_review_input_binding(self):
        # P2-12：放行凭证必须绑定写前审阅的 values/spec/policy 输入，
        # 未绑定输入的放行在生成与校验两侧都必须失败。
        with self.assertRaisesRegex(ValueError, "绑定"):
            contract_release.build_release([], {"status": "merged"}, str(self.main), str(self.fc),
                                           str(self.policy), None)
        release = contract_release.build_release([], {"status": "merged"}, str(self.main), str(self.fc),
                                                 str(self.policy), None, inputs=self.release_inputs())
        stale = {"schema": "dealer-ops/contract-review@1", "findings": [],
                 "llm_review": {"status": "merged"},
                 "release": {key: value for key, value in release.items() if key != "input_files"}}
        ok, problems = contract_release.verify_release(stale)
        self.assertFalse(ok)
        self.assertTrue(any("输入文件哈希绑定" in item for item in problems), problems)
        self.values.write_text('{"changed": true}', encoding="utf-8")
        document = {"schema": "dealer-ops/contract-review@1", "findings": [],
                    "llm_review": {"status": "merged"}, "release": release}
        ok, problems = contract_release.verify_release(document)
        self.assertFalse(ok)
        self.assertTrue(any("已变更" in item for item in problems), problems)

    def test_write_gate_requires_review_and_binds_high_resolutions(self):
        missing = subprocess.run([sys.executable, str(SCRIPTS / "apply_fill.py")],
                                 text=True, capture_output=True, timeout=10)
        self.assertNotEqual(missing.returncode, 0)
        self.assertIn("--review", missing.stderr)

        llm = self.root / "llm-high.json"
        llm.write_text(json.dumps({"findings": [{"severity": "high", "message": "需业务确认"}]}),
                       encoding="utf-8")
        out = self.root / "写前审阅.json"
        blocked = self.review_command(out, llm)
        self.assertNotEqual(blocked.returncode, 0)
        self.assertIn("尚无处理记录", blocked.stderr)
        resolutions = self.root / "resolutions.json"
        resolutions.write_text(json.dumps({"C-001": {"status": "接受风险", "note": "业务已确认", "owner": "kent"}}),
                               encoding="utf-8")
        approved = self.review_command(out, llm, resolutions=resolutions)
        self.assertEqual(approved.returncode, 0, approved.stdout + approved.stderr)
        document = json.loads(out.read_text(encoding="utf-8"))
        self.assertEqual(contract_release.check_write_review(document), [])
        document["findings"][0]["message"] = "审阅后篡改"
        self.assertTrue(any("已变更" in item for item in contract_release.check_write_review(document)))

    def test_write_gate_binds_actual_values_and_unchanged_inputs(self):
        llm = self.root / "llm.json"
        llm.write_text('{"findings": []}', encoding="utf-8")
        out = self.root / "写前审阅.json"
        run = self.review_command(out, llm)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        document = json.loads(out.read_text(encoding="utf-8"))
        other = self.root / "乙合同-values.json"
        other.write_text("{}", encoding="utf-8")
        problems = contract_release.check_write_review(
            document, {"spec": str(self.values), "values": str(other)}
        )
        self.assertTrue(any("不是写前审阅绑定" in item for item in problems), problems)
        self.values.write_text('{"changed": true}', encoding="utf-8")
        problems = contract_release.check_write_review(document)
        self.assertTrue(any("已变更" in item for item in problems), problems)

    def test_incremental_plan_binds_user_confirmed_values(self):
        # P2-11：增量（无指纹）计划不得绕过信任根——字段须在规格内、须有用户确认值、
        # 落笔文本须与确认值（或其渲染结果）一致。
        from apply_fill import check_incremental_plan
        spec = self.root / "spec.json"
        spec.write_text(json.dumps({
            "fields": [{"id": "JX-01", "label": "编号", "render": "主合同编号：{v}", "vtype": "text"},
                       {"id": "JX-02", "label": "名称", "render": "{v}", "vtype": "text"}],
            "cells": [], "grids": [],
        }), encoding="utf-8")
        values = self.root / "values-ok.json"
        values.write_text(json.dumps({"JX-01": "JX-2026-001"}), encoding="utf-8")
        plan = {"ops": [{"field_id": "JX-01", "kind": "replace", "find": "＿" * 12, "occurrence": 1,
                         "value": "主合同编号：JX-2026-001", "expect": "主合同编号：JX-2026-001"}]}
        self.assertEqual(check_incremental_plan(plan, str(spec), str(values)), [])
        tampered = json.loads(json.dumps(plan))
        tampered["ops"][0]["value"] = tampered["ops"][0]["expect"] = "主合同编号：HACK-001"
        self.assertTrue(check_incremental_plan(tampered, str(spec), str(values)))
        unconfirmed = json.loads(json.dumps(plan))
        unconfirmed["ops"][0]["field_id"] = "JX-02"
        problems = check_incremental_plan(unconfirmed, str(spec), str(values))
        self.assertTrue(any("确认值" in item for item in problems), problems)

    def test_final_gate_detects_main_attachment_and_policy_changes(self):
        for role, target in (("main", self.main), ("anti_diversion", self.fc), ("price_policy", self.policy)):
            with self.subTest(role=role):
                case = self.root / role
                case.mkdir()
                main, fc, policy = case / "主合同.docx", case / "附件.docx", case / "政策.xlsx"
                main.write_text("主合同", encoding="utf-8")
                fc.write_text("附件", encoding="utf-8")
                policy.write_bytes(self.policy.read_bytes())
                document = self.final_document(main, fc, policy)
                if role == "price_policy":
                    policy.write_bytes(b"changed")
                else:
                    target_in_case = main if role == "main" else fc
                    target_in_case.write_text("放行后换文", encoding="utf-8")
                ok, problems = contract_release.verify_final_release(document, str(main))
                self.assertFalse(ok)
                self.assertTrue(any("已变更" in item for item in problems), problems)

    def test_final_gate_binds_result_path_and_is_mandatory(self):
        document = self.final_document()
        ok, problems = contract_release.verify_final_release(document, str(self.main))
        self.assertTrue(ok, problems)
        other = self.root / "另一份主合同.docx"
        other.write_text(self.main.read_text(encoding="utf-8"), encoding="utf-8")
        ok, problems = contract_release.verify_final_release(document, str(other))
        self.assertFalse(ok)
        self.assertTrue(any("--result" in item for item in problems))
        missing = subprocess.run([sys.executable, str(SCRIPTS / "verify_fill.py")],
                                 text=True, capture_output=True, timeout=10)
        self.assertNotEqual(missing.returncode, 0)
        self.assertIn("--review-release", missing.stderr)


if __name__ == "__main__":
    unittest.main()
