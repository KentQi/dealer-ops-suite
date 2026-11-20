#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"
sys.path.insert(0, str(SCRIPTS))
import make_master
import make_todos


class TodoTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix=".test-todos-", dir=ROOT)
        self.root = Path(self.tmp.name)
        self.master = self.root / "_主档.xlsx"
        make_master.build(str(self.master))
        wb = openpyxl.load_workbook(self.master)
        contract = wb["合同"]
        contract["B2"] = "月结买断"
        contract["E2"] = "月结买断"
        contract["H2"] = "2026-12-05"  # 买断日
        contract["I2"] = "2026-12-05"  # 付款截止日，星期六；无顺延规则不得擅自改日期
        contract["L2"] = ""  # 清除模板中的示例顺延规则
        wb.save(self.master)

    def tearDown(self):
        self.tmp.cleanup()

    def test_no_calendar_marks_workdays_estimated_and_preserves_payment_date(self):
        items, _ = make_todos.generate(str(self.master), "202611")
        payment = next(item for item in items if item["type"] == "付款截止")
        self.assertEqual(payment["due_date"], "2026-12-05")
        self.assertFalse(payment["estimated"])
        workdays = [item for item in items if item["type"] in ("对账异议期", "对账单签回期")]
        self.assertTrue(workdays)
        self.assertTrue(all(item["estimated"] and item["due_date"] is None for item in workdays))

    def test_holiday_and_adjusted_workweek_calendar(self):
        calendar = {"years": {2026}, "holidays": {date(2026, 12, 3)},
                    "workdays": {date(2026, 12, 5)}}
        due, estimated = make_todos.workday_due(date(2026, 12, 1), 3, calendar)
        self.assertFalse(estimated)
        self.assertEqual(due, date(2026, 12, 5))  # 周六调休上班；12 月 3 日法定节假日不计
        adjusted, estimated, _ = make_todos.adjust_payment(date(2026, 12, 3), "法定节假日顺延", calendar)
        self.assertFalse(estimated)
        self.assertEqual(adjusted, date(2026, 12, 4))

    def test_reminder_layer_does_not_execute_contract_liability(self):
        items, _ = make_todos.generate(str(self.master), "202611")
        reminder = next(item for item in items if item["type"] == "履约提醒")
        self.assertIn("不自动", reminder["title"])
        self.assertTrue(reminder["evidence"])
        self.assertNotIn("返利金额", reminder)

    def test_cli_writes_machine_readable_todos(self):
        output = self.root / "待办.json"
        run = subprocess.run([sys.executable, str(SCRIPTS / "make_todos.py"), str(self.master),
                              "--period", "202611", "--out", str(output)], text=True, capture_output=True)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        result = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(result["schema"], "dealer-ops/todos@1")
        self.assertEqual(result["master"]["sha256"], hashlib.sha256(self.master.read_bytes()).hexdigest())


class FutureAndNoShiftTodoTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix=".test-future-todos-", dir=ROOT)
        self.master = Path(self.tmp.name) / "_主档.xlsx"
        make_master.build(str(self.master))

    def tearDown(self):
        self.tmp.cleanup()

    def test_no_shift_and_future_contract_filter(self):
        wb = openpyxl.load_workbook(self.master)
        ws = wb["合同"]
        ws["L2"] = "法定节假日不顺延"
        future = [cell.value for cell in ws[2]]
        future[0], future[2], future[3] = "HT-FUTURE", "2027-01-01", "2027-12-31"
        ws.append(future)
        mid_month = list(future)
        mid_month[0], mid_month[2], mid_month[3] = "HT-MID-NOV", "2026-11-15", "2026-12-31"
        ws.append(mid_month)
        wb.save(self.master)
        calendar = {"years": {2026, 2027}, "holidays": set(), "workdays": set()}
        items, _warnings = make_todos.generate(str(self.master), "202611", calendar)
        payment = next(item for item in items if item["type"] == "付款截止")
        self.assertEqual(payment["due_date"], "2026-12-05")
        numbers = {item["contract_no"] for item in items}
        self.assertNotIn("HT-FUTURE", numbers)
        self.assertIn("HT-MID-NOV", numbers)


class WritebackTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix=".test-writeback-", dir=ROOT)
        self.root = Path(self.tmp.name)
        self.master = self.root / "_主档.xlsx"
        make_master.build(str(self.master))
        self.payload = self.root / "_回写.json"
        self.payload_data = {
            "period": "202611", "contract_no": "HT-TEST", "business_type": "月结买断",
            "asof": "2026-11-30", "opening_qty": 0, "opening_amount": 0,
            "settle_qty": -5, "settle_amount": -250, "diff_qty": -2, "diff_amount": -110,
            "sticker_amount": -10, "sticker_pay": "待定", "buyout_date": "2026-12-05",
            "payment_due_date": "2026-12-15", "status": "待确认",
        }
        self.payload.write_text(json.dumps(self.payload_data, ensure_ascii=False), encoding="utf-8")
        self.workbook = self.root / "对账.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "01-对账单"
        wb.save(self.workbook)
        self.result = self.root / "_结算结果.json"
        self.result.write_text(json.dumps({
            "schema": "dealer-ops/settlement-result@1", "business_type": "monthly_buyout",
            "rule_result": {"business_type": "monthly_buyout"},
            "master": {"path": str(self.master), "sha256": hashlib.sha256(self.master.read_bytes()).hexdigest()},
            "workbook": {"path": str(self.workbook),
                         "sha256": hashlib.sha256(self.workbook.read_bytes()).hexdigest()},
            "writeback": self.payload_data, "writeback_path": str(self.payload),
        }, ensure_ascii=False), encoding="utf-8")
        self.packet = self.root / "packet.json"
        period = {"start": "2026-11-01", "end": "2026-11-30"}
        self.packet.write_text(json.dumps({
            "schema": "dealer-ops/review-packet@1", "period": period,
            "workbook": {"path": str(self.workbook),
                         "sha256": hashlib.sha256(self.workbook.read_bytes()).hexdigest()},
        }, ensure_ascii=False), encoding="utf-8")
        self.report = self.root / "report.json"
        self.report.write_text(json.dumps({
            "schema": "dealer-ops/review-report@1",
            "reviewer": {"name": "测试复核", "model": "test"},
            "reviewed_at": "2026-11-20T08:00:00+08:00",
            "workbook_sha256": hashlib.sha256(self.workbook.read_bytes()).hexdigest(),
            "verdict": "通过",
            "checklist": [{"id": f"R{index}", "result": "通过", "note": "已核对"} for index in range(1, 11)],
            "findings": [], "summary": "复核通过", "period": period,
        }, ensure_ascii=False), encoding="utf-8")

    def tearDown(self):
        self.tmp.cleanup()

    def run_writeback(self, write=False):
        cmd = [sys.executable, str(SCRIPTS / "writeback_ledger.py"), str(self.master),
               "--from-json", str(self.payload), "--json"] + (["--write"] if write else [])
        run = subprocess.run(cmd, text=True, capture_output=True, timeout=20)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        return json.loads(run.stdout.strip().splitlines()[-1])

    def apply_review(self):
        run = subprocess.run([sys.executable, str(SCRIPTS / "apply_review.py"), str(self.workbook),
                              str(self.report), "--packet", str(self.packet), "--apply"],
                             text=True, capture_output=True, timeout=20)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)

    def test_write_is_blocked_without_review_lock(self):
        before = self.master.read_bytes()
        run = subprocess.run([sys.executable, str(SCRIPTS / "writeback_ledger.py"), str(self.master),
                              "--from-json", str(self.payload), "--write", "--json"],
                             text=True, capture_output=True, timeout=20)
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("复核锁", run.stderr)
        self.assertEqual(self.master.read_bytes(), before)

    def test_signed_net_return_dry_run_write_and_idempotency(self):
        before = hashlib.sha256(self.master.read_bytes()).hexdigest()
        result = self.run_writeback()
        self.assertEqual(result["mode"], "dry-run")
        self.assertEqual(result["changes"]["settle_amount"]["new"], -250)
        self.assertEqual(hashlib.sha256(self.master.read_bytes()).hexdigest(), before)
        self.apply_review()
        result = self.run_writeback(write=True)
        self.assertEqual(result["mode"], "write")
        wb = openpyxl.load_workbook(self.master, data_only=True)
        headers = [cell.value for cell in wb["期间"][1]]
        row = next(row for row in wb["期间"].iter_rows(min_row=2, values_only=True) if row[0] == "202611")
        record = dict(zip(headers, row))
        self.assertEqual(record["结算金额"], -250)
        self.assertEqual(record["贴花金额"], -10)
        self.assertEqual(record["差异金额"], -110)
        self.assertEqual(record["买断日"], "2026-12-05")
        self.assertEqual(record["付款截止日"], "2026-12-15")
        self.assertEqual(self.run_writeback(write=True)["mode"], "noop")


if __name__ == "__main__":
    unittest.main()
