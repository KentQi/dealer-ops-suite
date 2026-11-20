#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"
sys.path.insert(0, str(SCRIPTS))
import check_policy
import contract_release

BAR = "690000000001"


class PolicyTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix=".test-policy-", dir=ROOT)
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def analyse(self, rows, headers=None):
        path = self.root / "最终价格政策.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "价格政策"
        ws.append(headers or ["商品条码", "商品名称", "结算价", "生效起", "生效止"])
        for row in rows:
            ws.append(row)
        wb.save(path)
        return check_policy.analyse(str(path))

    def codes(self, result):
        return {item["id"] for item in result["needs_confirmation"]}

    def test_two_legal_scd2_versions_are_not_duplicates(self):
        result = self.analyse([[BAR, "商品", 50, "2026-10-01", "2026-10-15"],
                               [BAR, "商品", 55, "2026-10-16", None]])
        self.assertEqual(result["needs_confirmation"], [])

    def test_overlap_and_invalid_value_are_blockers(self):
        result = self.analyse([[BAR, "商品", 50, "2026-10-01", "2026-10-16"],
                               [BAR, "商品", 55, "2026-10-16", None]])
        self.assertTrue(result["blocking_findings"])
        result = self.analyse([[BAR, "商品", "待定", "2026-10-01", None]])
        self.assertIn("P4", self.codes(result))
        self.assertTrue(result["blocking_findings"])

    def test_gap_is_high_and_missing_price_is_blocker(self):
        result = self.analyse([[BAR, "商品", 50, "2026-10-01", "2026-10-10"],
                               [BAR, "商品", 55, "2026-10-16", None]])
        self.assertIn("P5", self.codes(result))
        self.assertFalse(result["blocking_findings"])
        result = self.analyse([[BAR, "商品", None, "2026-10-01", None]])
        self.assertIn("P4", self.codes(result))
        self.assertTrue(result["blocking_findings"])

    def test_same_barcode_without_effective_date_needs_confirmation(self):
        result = self.analyse([[BAR, "商品", 50, None, None], [BAR, "商品", 55, None, None]])
        self.assertIn("P1", self.codes(result))
        self.assertTrue(result["blocking_findings"])

    def test_missing_barcode_with_valid_name_and_price_is_blocked(self):
        result = self.analyse([[None, "商品", 50, "2026-10-01", None]])
        self.assertEqual(result["data_rows"], 1)
        self.assertEqual(result["columns"]["name"], 2)
        missing = [item for item in result["blocking_findings"] if item["type"] == "条码为空"]
        self.assertEqual(len(missing), 1)
        self.assertEqual(missing[0]["name"], "商品")

    def test_one_header_cannot_map_to_two_fields(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["商品条码", "商品名称", "结算价", "生效起"])
        original = check_policy.COL_ALIASES["name"]
        check_policy.COL_ALIASES["name"] = ("商品",)
        try:
            with self.assertRaisesRegex(ValueError, "同时映射"):
                check_policy.map_columns(ws, 1)
        finally:
            check_policy.COL_ALIASES["name"] = original

    def test_multiple_price_columns_are_blocked_like_recon_engines(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["商品条码", "商品名称", "供货价", "结算价", "生效起", "生效止"])
        with self.assertRaisesRegex(check_policy.PolicyColumnError, "不得任选一列"):
            check_policy.map_columns(ws, 1)

    def test_footer_text_in_note_or_source_column_does_not_block(self):
        headers = ["商品条码", "商品名称", "结算价", "生效起", "生效止", "依据文件", "备注"]
        result = self.analyse([
            [BAR, "商品", 50, "2026-10-01", None, None, None],
            [None, None, None, None, None, "表尾版本说明", "累计与起订条件说明"],
        ], headers)
        self.assertEqual(result["data_rows"], 1)
        self.assertEqual(result["needs_confirmation"], [])


class ReleaseTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix=".test-release-", dir=ROOT)
        self.root = Path(self.tmp.name)
        for name in ("主合同.docx", "防窜货附件.docx"):
            (self.root / name).write_text(name, encoding="utf-8")
        self.policy = self.root / "政策.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "价格政策"
        for row in (["商品条码", "商品名称", "结算价", "生效起", "生效止"],
                    [BAR, "商品", 50, "2026-10-01", None]):
            ws.append(row)
        wb.save(self.policy)

    def tearDown(self):
        self.tmp.cleanup()

    def command(self, out, extra=()):
        values = self.root / "values.json"
        values.write_text("{}", encoding="utf-8")
        llm = self.root / "llm.json"
        llm.write_text('{"findings": []}', encoding="utf-8")
        fc_values = self.root / "fc-values.json"
        fc_values.write_text('{"FC-C1": "2026-10-01", "FC-C2": "2026-10-01"}', encoding="utf-8")
        cmd = [sys.executable, str(SCRIPTS / "review_contract_package.py"),
               "--spec", str(values), "--values", str(values),
               "--fc-spec", str(values), "--fc-values", str(fc_values),
               "--policy", str(self.policy), "--llm-findings", str(llm), "--out", str(out)]
        return subprocess.run(cmd + list(extra), text=True, capture_output=True, timeout=20)

    def release_args(self):
        return ["--release", "--main-file", str(self.root / "主合同.docx"),
                "--fc-file", str(self.root / "防窜货附件.docx")]

    def test_release_and_hash_verification(self):
        out = self.root / "审阅.json"
        run = self.command(out, self.release_args())
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        verify = subprocess.run([sys.executable, str(SCRIPTS / "review_contract_package.py"),
                                 "--verify-release", str(out)], text=True, capture_output=True)
        self.assertEqual(verify.returncode, 0, verify.stdout + verify.stderr)
        (self.root / "主合同.docx").write_text("放行后换文", encoding="utf-8")
        verify = subprocess.run([sys.executable, str(SCRIPTS / "review_contract_package.py"),
                                 "--verify-release", str(out)], text=True, capture_output=True)
        self.assertNotEqual(verify.returncode, 0)
        self.assertIn("已变更", verify.stderr)

    def test_release_verification_rejects_changed_review_inputs(self):
        out = self.root / "审阅.json"
        run = self.command(out, self.release_args())
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        (self.root / "values.json").write_text('{"changed": true}', encoding="utf-8")
        verify = subprocess.run([sys.executable, str(SCRIPTS / "review_contract_package.py"),
                                 "--verify-release", str(out)], text=True, capture_output=True)
        self.assertNotEqual(verify.returncode, 0)
        self.assertIn("已变更", verify.stderr)

    def test_missing_llm_blocks_release(self):
        values = self.root / "values.json"
        values.write_text("{}", encoding="utf-8")
        run = subprocess.run([sys.executable, str(SCRIPTS / "review_contract_package.py"),
                              "--spec", str(values), "--values", str(values),
                              "--fc-spec", str(values), "--fc-values", str(values),
                              "--policy", str(self.policy), "--out", str(self.root / "审阅.json"),
                              "--release", "--main-file", str(self.root / "主合同.docx"),
                              "--fc-file", str(self.root / "防窜货附件.docx")], text=True, capture_output=True)
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("LLM", run.stderr)

    def test_high_requires_exact_id_owner_and_note(self):
        findings = [{"id": "C-001", "severity": "high", "message": "需确认"}]
        with self.assertRaises(ValueError):
            contract_release.build_release(findings, {"status": "merged"}, str(self.root / "主合同.docx"),
                                           str(self.root / "防窜货附件.docx"), str(self.policy), None)

    def test_policy_overlap_cannot_be_released(self):
        wb = openpyxl.load_workbook(self.policy)
        wb.active.append([BAR, "商品", 55, "2026-10-01", None])
        wb.save(self.policy)
        run = self.command(self.root / "审阅.json", self.release_args())
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("阻断", run.stderr)

    def test_normal_review_keeps_policy_blocker_and_exits_2(self):
        # P2-7：普通审阅不得把政策 blocker 降级成 high；必须显示阻断并以 2 退出
        wb = openpyxl.load_workbook(self.policy)
        wb.active.append([BAR, "商品", 55, "2026-10-01", None])
        wb.save(self.policy)
        out = self.root / "审阅.json"
        run = self.command(out)
        self.assertEqual(run.returncode, 2, run.stdout + run.stderr)
        result = json.loads(out.read_text(encoding="utf-8"))
        self.assertGreaterEqual(result["summary"]["blocker"], 1)
        self.assertIn("阻断", run.stdout)


if __name__ == "__main__":
    unittest.main()
