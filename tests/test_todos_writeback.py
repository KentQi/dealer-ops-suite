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
import writeback_ledger


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

    def test_same_contract_and_kind_with_different_key_values_have_unique_ids(self):
        wb = openpyxl.load_workbook(self.master)
        ws = wb["履约条款"]
        contract_no = ws["A2"].value
        ws.append([contract_no, "履约指标", "季度3%", "季度指标为3%", "主合同.docx", "季度条款", "待确认", ""])
        ws.append([contract_no, "履约指标", "年度5%", "年度指标为5%", "主合同.docx", "年度条款", "待确认", ""])
        wb.save(self.master)
        items, _warnings = make_todos.generate(str(self.master), "202611")
        selected = [item for item in items if item.get("key_value") in ("季度3%", "年度5%")]
        self.assertEqual(len(selected), 2)
        self.assertEqual(len({item["id"] for item in selected}), 2)
        self.assertEqual(len({item["id"] for item in items}), len(items))

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


class PayTextAndDateParsingTest(unittest.TestCase):
    """P2-1／P2-4／P2-9：付款日文本、斜杠日期、跨年顺延的静默丢失与伪精确。"""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix=".test-paytext-", dir=ROOT)
        self.master = Path(self.tmp.name) / "_主档.xlsx"
        make_master.build(str(self.master))

    def tearDown(self):
        self.tmp.cleanup()

    def _contract_ws(self):
        wb = openpyxl.load_workbook(self.master)
        return wb, wb["合同"]

    def test_unparseable_pay_text_warns_and_makes_missing_todo(self):
        wb, ws = self._contract_ws()
        ws["I2"] = ""  # 清除付款截止日
        ws["G2"] = "月结30天"  # 付款日：无法解析出明确日
        wb.save(self.master)
        items, warnings = make_todos.generate(str(self.master), "202611")
        self.assertFalse(any(item["type"] == "付款截止" for item in items))
        missing = [item for item in items if item["type"] == "资料缺失"]
        self.assertEqual(len(missing), 1)
        self.assertTrue(any("月结30天" in warning for warning in warnings))

    def test_pay_text_day_over_28_is_not_presumed(self):
        wb, ws = self._contract_ws()
        ws["I2"] = ""
        ws["G2"] = "账期30日"
        wb.save(self.master)
        items, warnings = make_todos.generate(str(self.master), "202611")
        self.assertFalse(any(item["type"] == "付款截止" for item in items))
        self.assertEqual(len([item for item in items if item["type"] == "资料缺失"]), 1)
        self.assertTrue(any("账期30日" in warning for warning in warnings))

    def test_monthly_keyword_normalizes_business_type(self):
        wb, ws = self._contract_ws()
        ws["B2"] = "月结"  # 非完整「月结买断」也应识别为月结口径
        ws["E2"] = "月结"
        ws["I2"] = ""  # 无付款截止日、无付款日文本
        ws["G2"] = ""
        wb.save(self.master)
        items, warnings = make_todos.generate(str(self.master), "202611")
        self.assertEqual(len([item for item in items if item["type"] == "资料缺失"]), 1)
        self.assertFalse(any(item["type"] == "付款截止" for item in items))

    def test_slash_dates_are_parsed_and_invalid_dates_warn(self):
        wb, ws = self._contract_ws()
        future = [cell.value for cell in ws[2]]
        future[0], future[2], future[3] = "HT-SLASH", "2027/01/01", "2027/12/31"
        ws.append(future)
        bad = list(future)
        bad[0], bad[2] = "HT-BAD", "不是日期"
        ws.append(bad)
        wb.save(self.master)
        items, warnings = make_todos.generate(str(self.master), "202611")
        numbers = {item["contract_no"] for item in items}
        self.assertNotIn("HT-SLASH", numbers)  # 斜杠起始日被正确解析并过滤
        self.assertTrue(any("HT-BAD" in warning and "不是日期" in warning for warning in warnings))

    def test_slash_end_date_generates_expiry_todo(self):
        wb, ws = self._contract_ws()
        ws["D2"] = "2026/12/31"  # 终止日用斜杠，应生成到期提醒
        wb.save(self.master)
        items, warnings = make_todos.generate(str(self.master), "202611")
        expiry = [item for item in items if item["type"] == "合同到期"]
        self.assertEqual(len(expiry), 1)
        self.assertEqual(expiry[0]["due_date"], "2026-12-31")
        self.assertFalse(warnings)

    def test_adjust_payment_crossing_uncovered_year_is_estimated(self):
        calendar = {"years": {2026}, "holidays": {date(2026, 12, 28), date(2026, 12, 29),
                                                date(2026, 12, 30), date(2026, 12, 31)},
                    "workdays": set()}
        adjusted, estimated, note = make_todos.adjust_payment(date(2026, 12, 26), "法定节假日顺延", calendar)
        self.assertTrue(estimated)
        self.assertIsNone(adjusted)
        self.assertIn("日历", note)
        full = {"years": {2026, 2027}, "holidays": calendar["holidays"], "workdays": set()}
        adjusted, estimated, _note = make_todos.adjust_payment(date(2026, 12, 26), "法定节假日顺延", full)
        self.assertFalse(estimated)
        self.assertEqual(adjusted, date(2027, 1, 1))


class WorkdayConsistencyTest(unittest.TestCase):
    """P2-3：make_todos 与 ledger_read 工作日口径一致。"""

    def test_calendar_aware_workdays_match_between_scripts(self):
        calendar = {"years": {2026}, "holidays": {date(2026, 12, 3)},
                    "workdays": {date(2026, 12, 5)}}
        import ledger_read
        start = date(2026, 12, 1)
        due_todos, estimated_todos = make_todos.workday_due(start, 3, calendar)
        due_ledger = ledger_read.add_workdays(start, 3, calendar)
        self.assertFalse(estimated_todos)
        self.assertEqual(due_todos, due_ledger)
        self.assertEqual(due_ledger, date(2026, 12, 5))

    def test_weekend_only_basis_is_annotated_not_silent(self):
        import ledger_read
        import workdays
        start = date(2026, 10, 1)
        due, estimated, note = workdays.add_workdays(start, 5, None)
        self.assertEqual(due, ledger_read.add_workdays(start, 5, None))
        self.assertTrue(estimated)
        self.assertIn("仅按周末", note)
        self.assertIn("不含法定节假日", ledger_read.WORKDAY_BASIS_WEEKEND_ONLY)
        # 无日历时 make_todos 不生成精确日期，与 ledger_read 的周末口径语义明确分开
        due_todos, estimated_todos = make_todos.workday_due(start, 5, None)
        self.assertIsNone(due_todos)
        self.assertTrue(estimated_todos)

    def test_empty_calendar_matches_between_scripts(self):
        import ledger_read
        start = date(2026, 10, 1)
        due_calendar = {"years": {2026}, "holidays": set(), "workdays": set()}
        due_todos, estimated_cal = make_todos.workday_due(start, 5, due_calendar)
        self.assertFalse(estimated_cal)
        self.assertEqual(due_todos, ledger_read.add_workdays(start, 5, due_calendar))



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

    def test_result_writeback_requires_contract_business_and_amount(self):
        self.apply_review()
        original_result = json.loads(self.result.read_text(encoding="utf-8"))
        original_payload = self.payload.read_text(encoding="utf-8")
        try:
            for key in ("contract_no", "business_type", "settle_amount"):
                with self.subTest(missing=key):
                    document = json.loads(json.dumps(original_result))
                    document["writeback"].pop(key, None)
                    self.result.write_text(json.dumps(document, ensure_ascii=False), encoding="utf-8")
                    with self.assertRaisesRegex(ValueError, "必填字段"):
                        writeback_ledger.verify_write_provenance(
                            str(self.master), str(self.payload), str(self.result),
                            str(self.root / "_复核记录"), self.payload_data)
            zero_values = dict(self.payload_data, settle_amount=0)
            document = json.loads(json.dumps(original_result))
            document["writeback"]["settle_amount"] = 0
            self.result.write_text(json.dumps(document, ensure_ascii=False), encoding="utf-8")
            self.payload.write_text(json.dumps(zero_values, ensure_ascii=False), encoding="utf-8")
            writeback_ledger.verify_write_provenance(
                str(self.master), str(self.payload), str(self.result),
                str(self.root / "_复核记录"), zero_values)
        finally:
            self.result.write_text(json.dumps(original_result, ensure_ascii=False), encoding="utf-8")
            self.payload.write_text(original_payload, encoding="utf-8")

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
