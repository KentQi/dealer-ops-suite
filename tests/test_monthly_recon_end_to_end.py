#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""月结买断 Excel 端到端验收。"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"
sys.path.insert(0, str(SCRIPTS))
try:
    from tests.test_recon_end_to_end import BAR, workbook  # 作为包内测试运行
except ImportError:  # 按单文件直接运行时 tests 不是包
    from test_recon_end_to_end import BAR, workbook


class MonthlyBuyoutWorkbookTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(prefix=".test-monthly-", dir=ROOT)
        self.base = Path(self._tmp.name)
        self.master = self.base / "_主档.xlsx"
        self.receipts = self.base / "签收.xlsx"
        self.returns = self.base / "退货.xlsx"
        self.sales = self.base / "动销.xlsx"
        self.out = self.base / "202610月结买断对账.xlsx"
        workbook(self.master, {
            "SKU": [["商品条码", "标准品名", "结算属性"], [BAR, "脱敏商品", "结算类"]],
            "价格": [["商品条码", "结算价", "生效起", "生效止", "依据文件"],
                    [BAR, 50, datetime(2026, 10, 1), datetime(2026, 10, 15), "合同附件v1"],
                    [BAR, 55, datetime(2026, 10, 16), None, "调价政策v2"]],
            "贴花": [["商品条码", "贴花标准", "生效起", "生效止", "依据文件"],
                    [BAR, 2, datetime(2026, 10, 1), datetime(2026, 10, 15), "合同附件v1"],
                    [BAR, 3, datetime(2026, 10, 16), None, "调价政策v2"]],
            "期间": [["期间", "数据基准日", "期初数量", "期初金额", "结算数量", "结算金额",
                     "差异数量", "差异金额", "贴花金额", "贴花兑付方式", "状态", "备注"]],
        })
        workbook(self.receipts, {"签收": [
            ["流水号", "签收日期", "商品条码", "签收数量"],
            ["P1", "2026-10-10", BAR, 40], ["P2", "2026-10-20", BAR, 20],
        ]})
        workbook(self.returns, {"退货": [[
            "流水号", "退货日期", "商品条码", "退货数量", "原签收流水号", "原销售流水号",
            "原签收单价", "原销售已计提贴花", "原销售贴花标准",
        ]]})
        workbook(self.sales, {"动销": [
            ["流水号", "动销日期", "商品条码", "动销数量"],
            ["S1", "2026-10-12", BAR, 8], ["S2", "2026-10-22", BAR, 7],
        ]})

    def tearDown(self):
        self._tmp.cleanup()

    def config(self):
        return {
            "business_type": "monthly_buyout",
            "party_a": {"name": "脱敏甲方"}, "party_b": {"name": "脱敏乙方"},
            "contract": {"no": "TEST-M-2026", "method": "月结买断",
                         "buyout_date": "2026-11-05", "payment_due_date": "2026-11-15"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "base_dir": str(self.base), "master_path": str(self.master),
            "sources": {"receipts": str(self.receipts), "returns": str(self.returns), "sales": str(self.sales)},
            "out_file": str(self.out), "overwrite": True,
        }

    def run_engine(self, cfg=None, expect=0):
        env = dict(os.environ, RECON_CFG=json.dumps(cfg or self.config(), ensure_ascii=False))
        run = subprocess.run([sys.executable, str(SCRIPTS / "monthly_recon_engine.py")], env=env,
                             cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(run.returncode, expect, run.stdout + "\n" + run.stderr)
        if expect:
            return run, None, None
        with open(self.base / "_结算结果.json", encoding="utf-8") as fh:
            result = json.load(fh)
        with open(self.base / "_回写.json", encoding="utf-8") as fh:
            writeback = json.load(fh)
        return run, result, writeback

    def test_master_sheet_mapping_is_honored(self):
        # P2-5：_数据映射.json 的 master.sheets（实际表名 → 规范表名）在月结链路与铺货一致生效
        wb = openpyxl.load_workbook(self.master)
        for old, new in (("价格", "价格表"), ("贴花", "贴花表"), ("SKU", "商品档案")):
            wb[old].title = new
        wb.save(self.master)
        cfg = self.config()
        cfg["mapping"] = {"schema": "dealer-ops/data-mapping@1",
                          "master": {"sheets": {"价格表": "价格", "贴花表": "贴花", "商品档案": "SKU"}}}
        _run, result, writeback = self.run_engine(cfg=cfg)
        self.assertEqual(result["rule_result"]["buyout_qty"], "60")
        self.assertEqual(result["rule_result"]["buyout_amount"], "3100.00")
        self.assertEqual(writeback["settle_amount"], 3100)

    def test_decoy_shipment_quantity_does_not_override_receipt_quantity(self):
        wb = openpyxl.load_workbook(self.receipts)
        ws = wb["签收"]
        ws.insert_cols(4)
        ws.cell(1, 4).value = "发货数量"
        for row in range(2, ws.max_row + 1):
            ws.cell(row, 4).value = 999
        wb.save(self.receipts)
        _run, result, writeback = self.run_engine()
        self.assertEqual(result["rule_result"]["buyout_qty"], "60")
        self.assertEqual(result["rule_result"]["buyout_amount"], "3100.00")
        self.assertEqual(writeback["settle_qty"], 60)

    def test_receipt_buyout_and_sales_decal_are_independent(self):
        _run, result, writeback = self.run_engine()
        rule = result["rule_result"]
        self.assertEqual(rule["buyout_qty"], "60")
        self.assertEqual(rule["buyout_amount"], "3100.00")
        self.assertEqual(rule["sales_qty_for_decal_only"], "15")
        self.assertEqual(rule["decal_amount"], "37.00")
        self.assertEqual(rule["buyout_date"], "2026-11-05")
        self.assertEqual(rule["payment_due_date"], "2026-11-15")
        self.assertEqual(writeback["settle_qty"], 60)
        self.assertEqual(writeback["settle_amount"], 3100)
        self.assertEqual(writeback["diff_qty"], 0)
        wb = openpyxl.load_workbook(self.out, data_only=True)
        self.assertEqual(wb.sheetnames, ["00-对账说明", "01-月结对账单", "02-签收与退货", "03-乙方动销贴花"])
        self.assertEqual(wb["01-月结对账单"]["E5"].value, 60)
        self.assertEqual(wb["01-月结对账单"]["F5"].value, 15)
        check = subprocess.run([sys.executable, str(SCRIPTS / "writeback_ledger.py"), str(self.master),
                                "--from-json", str(self.base / "_回写.json"), "--json"],
                               text=True, capture_output=True, timeout=15)
        self.assertEqual(check.returncode, 0, check.stdout + "\n" + check.stderr)

    def test_return_reduces_receipt_at_original_price(self):
        wb = openpyxl.load_workbook(self.returns)
        wb["退货"].append(["R1", "2026-10-26", BAR, 5, "P1", "S1", None, "是", None])
        wb.save(self.returns)
        _run, result, writeback = self.run_engine()
        rule = result["rule_result"]
        self.assertEqual(rule["buyout_qty"], "55")
        self.assertEqual(rule["return_amount"], "-250.00")
        self.assertEqual(rule["buyout_amount"], "2850.00")
        self.assertEqual(rule["decal_return_amount"], "-10.00")
        self.assertEqual(rule["decal_amount"], "27.00")
        self.assertEqual(writeback["settle_qty"], 55)
        wb = openpyxl.load_workbook(self.out, data_only=True)
        self.assertEqual(wb["02-签收与退货"]["G7"].value, 50)
        self.assertEqual(wb["02-签收与退货"]["H7"].value, -250)

    def test_non_settlement_rows_are_excluded_and_unknown_attributes_block(self):
        other = "690000000002"
        wb = openpyxl.load_workbook(self.master)
        wb["SKU"].append([other, "非结算赠品", "非结算类"])
        wb.save(self.master)
        wb = openpyxl.load_workbook(self.receipts)
        wb["签收"].append(["P3", "2026-10-21", other, 99])
        wb.save(self.receipts)
        wb = openpyxl.load_workbook(self.sales)
        wb["动销"].append(["S3", "2026-10-22", other, 99])
        wb.save(self.sales)
        _run, result, _writeback = self.run_engine()
        self.assertEqual(result["rule_result"]["buyout_amount"], "3100.00")
        self.assertEqual(result["rule_result"]["decal_amount"], "37.00")
        self.assertEqual(result["rule_result"]["ignored_non_settlement_barcodes"], [other])
        wb = openpyxl.load_workbook(self.master)
        wb["SKU"]["C3"] = "待确认"
        wb.save(self.master)
        run, _result, _writeback = self.run_engine(expect=1)
        self.assertIn("未明确登记结算类", run.stderr)

    def test_duplicate_return_transaction_ids_are_blocked(self):
        wb = openpyxl.load_workbook(self.returns)
        for _ in range(2):
            wb["退货"].append(["R1", "2026-10-26", BAR, 1, "P1", "S1", None, "是", None])
        wb.save(self.returns)
        run, _result, _writeback = self.run_engine(expect=1)
        self.assertIn("DUPLICATE_TRANSACTION_ID", run.stderr)

    def test_monthly_review_packet_has_no_opening_and_keeps_detail(self):
        cfg = self.config()
        _run, result, _writeback = self.run_engine(cfg)
        sys.path.insert(0, str(SCRIPTS))
        from make_review_packet import build_packet
        previous = os.environ.get("RECON_CFG")
        os.environ["RECON_CFG"] = json.dumps(cfg, ensure_ascii=False)
        try:
            packet = build_packet(str(self.out), str(self.master))
        finally:
            if previous is None:
                os.environ.pop("RECON_CFG", None)
            else:
                os.environ["RECON_CFG"] = previous
        self.assertEqual(packet["business_type"], "monthly_buyout")
        self.assertIsNone(packet["figures"]["opening_qty"])
        self.assertIsNone(packet["figures"]["diff_qty"])
        self.assertEqual(packet["figures"]["receipt_qty"], 60)
        self.assertEqual(packet["figures"]["settled_amount"], 3100)
        self.assertEqual(packet["figures"]["buyout_date"], "2026-11-05")
        self.assertEqual(len(packet["priced_rows"]), 2)
        self.assertEqual(len(packet["tiehua_rows"]), 2)
        self.assertEqual(packet["sku_rows"][0]["sales_qty_for_decal_only"], 15)
        rule_ref = packet["rule_result"]
        self.assertEqual(set(rule_ref), {"_external", "sha256", "rows", "business_type"})
        self.assertEqual(rule_ref["_external"], "202610月结买断对账_复核包_rule_result.json")
        self.assertEqual(rule_ref["rows"], 4)  # 签收 2 + 退货 0 + 贴花动销 2
        self.assertEqual(rule_ref["business_type"], "monthly_buyout")
        sidecar = self.base / "_复核记录" / rule_ref["_external"]
        with open(sidecar, encoding="utf-8") as fh:
            self.assertEqual(json.load(fh), result["rule_result"])
        self.assertEqual(rule_ref["sha256"],
                         __import__("hashlib").sha256(sidecar.read_bytes()).hexdigest())
        self.assertEqual({row["role"] for row in packet["sources"]}, {"receipts", "returns", "sales"})

    def test_multiple_price_columns_are_blocked(self):
        wb = openpyxl.load_workbook(self.master)
        ws = wb["价格"]
        ws.cell(1, 6).value = "供货价"
        for row in range(2, ws.max_row + 1):
            ws.cell(row, 6).value = 999
        wb.save(self.master)
        run, _result, _writeback = self.run_engine(expect=1)
        self.assertIn("AMBIGUOUS_RATE_COLUMN", run.stderr)
        self.assertIn("不得任选一列", run.stderr)

    def test_statement_no_is_preserved_in_both_writeback_payloads(self):
        cfg = self.config()
        cfg["statement_no"] = "DJ-M-202610-01"
        _run, result, writeback = self.run_engine(cfg)
        self.assertEqual(writeback["statement_no"], "DJ-M-202610-01")
        self.assertEqual(result["writeback"], writeback)
        self.assertEqual(writeback["contract_no"], "TEST-M-2026")
        self.assertEqual(writeback["business_type"], "月结买断")

    def test_missing_dates_and_original_receipt_are_blocked(self):
        cfg = self.config()
        del cfg["contract"]["payment_due_date"]
        run, _result, _writeback = self.run_engine(cfg, expect=1)
        self.assertIn("付款截止日", run.stderr)
        wb = openpyxl.load_workbook(self.returns)
        wb["退货"].append(["R1", "2026-10-26", BAR, 5, None, "S1", None, "是", None])
        wb.save(self.returns)
        run, _result, _writeback = self.run_engine(expect=1)
        self.assertIn("ORIGINAL_PRICE_REQUIRED", run.stderr)


if __name__ == "__main__":
    unittest.main()
