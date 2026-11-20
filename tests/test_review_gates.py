#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"


class ReviewGateTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix=".test-review-gates-", dir=ROOT)
        self.root = Path(self.tmp.name)
        self.workbook = self.root / "对账.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "01-对账单"
        wb.save(self.workbook)
        self.pre_sha = hashlib.sha256(self.workbook.read_bytes()).hexdigest()
        self.period = {"start": "2026-10-01", "end": "2026-10-31"}
        self.packet = self.root / "packet.json"
        self.packet.write_text(json.dumps({
            "schema": "dealer-ops/review-packet@1", "period": self.period,
            "workbook": {"path": str(self.workbook), "sha256": self.pre_sha},
        }, ensure_ascii=False), encoding="utf-8")
        self.report_data = {
            "schema": "dealer-ops/review-report@1",
            "reviewer": {"name": "测试复核", "model": "test"},
            "reviewed_at": "2026-11-20T08:00:00+08:00",
            "workbook_sha256": self.pre_sha, "verdict": "通过",
            "checklist": [{"id": f"R{index}", "result": "通过", "note": "已核对"}
                          for index in range(1, 11)],
            "findings": [], "summary": "复核通过", "period": self.period,
        }
        self.report = self.root / "report.json"
        self.save_report()

    def tearDown(self):
        self.tmp.cleanup()

    def save_report(self):
        self.report.write_text(json.dumps(self.report_data, ensure_ascii=False), encoding="utf-8")

    def run_review(self, *extra):
        return subprocess.run([sys.executable, str(SCRIPTS / "apply_review.py"), str(self.workbook),
                               str(self.report), *extra], text=True, capture_output=True, timeout=20)

    def test_apply_requires_packet(self):
        run = self.run_review("--apply")
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("必须提供 --packet", run.stderr)
        self.assertFalse((self.root / "_复核记录" / "复核锁.json").exists())

    def test_failed_checklist_cannot_have_pass_verdict(self):
        for item in self.report_data["checklist"]:
            item["result"] = "不通过"
        self.save_report()
        run = self.run_review("--packet", str(self.packet), "--apply")
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("不得为「通过」", run.stderr)
        self.assertFalse((self.root / "_复核记录" / "复核锁.json").exists())

    def test_valid_packet_applies_and_checks_lock(self):
        run = self.run_review("--packet", str(self.packet), "--apply")
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        lock_path = self.root / "_复核记录" / "复核锁.json"
        lock = json.loads(lock_path.read_text(encoding="utf-8"))
        self.assertEqual(lock["pre_sha256"], self.pre_sha)
        self.assertEqual(lock["post_sha256"], hashlib.sha256(self.workbook.read_bytes()).hexdigest())
        check = subprocess.run([sys.executable, str(SCRIPTS / "apply_review.py"), str(self.workbook),
                                "--check"], text=True, capture_output=True, timeout=20)
        self.assertEqual(check.returncode, 0, check.stdout + check.stderr)


if __name__ == "__main__":
    unittest.main()
