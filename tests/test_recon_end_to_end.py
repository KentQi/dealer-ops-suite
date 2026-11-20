#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""铺货代销 Excel 端到端脱敏验收：真实出账，不用 mock 替代规则核心。"""

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
BAR = "690000000001"


def workbook(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


class ConsignmentWorkbookTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(prefix=".test-recon-", dir=ROOT)
        self.base = Path(self._tmp.name)
        self.master = self.base / "_主档.xlsx"
        self.jia = self.base / "BIP期初.xlsx"
        self.sales = self.base / "乙方销售.xlsx"
        self.inventory = self.base / "乙方库存.xlsx"
        self.out = self.base / "202610铺货代销对账.xlsx"
        workbook(self.master, {
            "SKU": [["商品条码", "标准品名", "别名", "规格", "单位", "品牌", "结算属性"],
                    [BAR, "脱敏商品", "脱敏商品", "10ml", "件", "测试", "结算类"]],
            "价格": [["商品条码", "结算价", "生效起", "生效止", "依据文件"],
                    [BAR, 50, datetime(2026, 10, 1), datetime(2026, 10, 15), "合同附件v1"],
                    [BAR, 55, datetime(2026, 10, 16), None, "调价政策v2"]],
            "贴花": [["商品条码", "贴花标准", "生效起", "生效止", "依据文件"],
                    [BAR, 2, datetime(2026, 10, 1), datetime(2026, 10, 15), "合同附件v1"],
                    [BAR, 3, datetime(2026, 10, 16), None, "调价政策v2"]],
            "期间": [["期间", "数据基准日", "期初数量", "期初金额", "结算数量", "结算金额",
                     "差异数量", "差异金额", "贴花金额", "贴花兑付方式", "状态", "备注"]],
        })
        workbook(self.jia, {
            "BIP挂账": [["商品名称", "未结算数量", "初始数量", "物料属性", "物料编码"],
                        ["脱敏商品", 100, 100, "正装", ""]],
        })
        workbook(self.inventory, {
            "库存": [["商品名称", "条码", "规格", "分类", "库存数量", "销售价", "品牌"],
                     ["脱敏商品", BAR, "10ml", "测试", 37, 100, "测试"]],
        })
        self.rows = [
            ["流水号", "销售时间", "门店", "商品名称", "条码", "规格", "单位", "销售数量", "品牌",
             "业务类型", "原销售流水号", "原销售单价", "原销售已计提贴花", "原销售贴花标准"],
            ["S1", "2026-10-10", "一店", "脱敏商品", BAR, "10ml", "件", 40, "测试", "销售", None, None, None, None],
            ["S2", "2026-10-20", "一店", "脱敏商品", BAR, "10ml", "件", 20, "测试", "销售", None, None, None, None],
        ]

    def tearDown(self):
        self._tmp.cleanup()

    def run_engine(self, *, sales_rows=None, inventory_qty=37, authorization=None,
                   statement_no=None, cfg_out=None, expect=0):
        workbook(self.sales, {"10月销售": sales_rows or self.rows})
        wb = openpyxl.load_workbook(self.inventory)
        wb["库存"]["E2"] = inventory_qty
        wb.save(self.inventory)
        cfg = {
            "business_type": "consignment",
            "party_a": {"name": "脱敏甲方", "tax": "TEST-A"},
            "party_b": {"name": "脱敏乙方", "tax": "TEST-B"},
            "bank": {"account_name": "脱敏甲方", "account_no": "000001", "bank_name": "测试银行"},
            "contract": {"no": "TEST-2026", "method": "铺货代销", "basis": "实销实结"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "base_dir": str(self.base), "master_path": str(self.master),
            "sources": {"jia": self.jia.name, "sales": self.sales.name, "inventory": self.inventory.name},
            "out_file": str(self.out), "overwrite": True, "v17": {"enabled": False},
        }
        if authorization is not None:
            cfg["variance_authorization"] = authorization
        if statement_no is not None:
            cfg["statement_no"] = statement_no
        if cfg_out is not None:
            Path(cfg_out).write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        env = dict(os.environ, RECON_CFG=json.dumps(cfg, ensure_ascii=False))
        run = subprocess.run([sys.executable, str(SCRIPTS / "recon_engine.py")], env=env,
                             cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(run.returncode, expect, run.stdout + "\n" + run.stderr)
        if expect:
            return None, None, run
        self.assertTrue(self.out.exists())
        with open(self.base / "_结算结果.json", encoding="utf-8") as fh:
            result = json.load(fh)
        with open(self.base / "_回写.json", encoding="utf-8") as fh:
            writeback = json.load(fh)
        return result, writeback, run

    def test_decoy_shipment_quantity_does_not_override_sales_quantity(self):
        rows = [list(row) for row in self.rows]
        for index, row in enumerate(rows):
            row.insert(7, "发货数量" if index == 0 else 999)
        result, writeback, _run = self.run_engine(sales_rows=rows)
        self.assertEqual(result["rule_result"]["quantity_reconciliation"][0]["net_sales_qty"], "60")
        self.assertEqual(result["rule_result"]["gross_sales_amount"], "3100.00")
        self.assertEqual(writeback["settle_qty"], 60)

    def test_fractional_quantities_are_not_truncated(self):
        """N4 回归：源表实数数量（4.7／95.3）不得在输入端 int 取整。
        期初 100、销售 4.7、库存 95.3、单价 50 → 结算 235.00，盈亏 0。"""
        rows = [list(row) for row in self.rows]
        rows[1][7] = 4.7
        rows[2] = None  # 只保留一笔销售
        rows = [row for row in rows if row is not None]
        result, writeback, run = self.run_engine(sales_rows=rows, inventory_qty=95.3)
        rule = result["rule_result"]["quantity_reconciliation"][0]
        self.assertEqual(rule["net_sales_qty"], "4.7")
        self.assertEqual(rule["variance_qty"], "0")
        self.assertEqual(result["rule_result"]["sales_settlement_amount"], "235.00")
        self.assertEqual(writeback["settle_qty"], 4.7)
        self.assertIn("净动销 4.7 件", run.stderr)

    def test_sales_only_unknown_barcode_is_blocked(self):
        extra = list(self.rows[1])
        extra[0], extra[4] = "S3", "690000000099"
        _result, _writeback, run = self.run_engine(sales_rows=self.rows + [extra], expect=1)
        self.assertIn("主档未建档", run.stderr)

    def test_blank_required_quantities_are_blocked(self):
        workbook(self.sales, {"10月销售": self.rows})
        originals = {path: path.read_bytes() for path in (self.jia, self.sales, self.inventory)}
        blank_sales = [list(row) for row in self.rows]
        blank_sales[1][7] = None
        cases = ((self.jia, "BIP挂账", "B2", {}),
                 (self.sales, "10月销售", "H2", {"sales_rows": blank_sales}),
                 (self.inventory, "库存", "E2", {"inventory_qty": None}))
        for path, sheet, cell, kwargs in cases:
            with self.subTest(source=path.name):
                wb = openpyxl.load_workbook(path)
                wb[sheet][cell] = None
                wb.save(path)
                _result, _writeback, run = self.run_engine(expect=1, **kwargs)
                self.assertIn("为空", run.stderr)
                for original, content in originals.items():
                    original.write_bytes(content)

    def test_review_packet_rejects_changed_sources_or_master(self):
        self.run_engine()
        from make_review_packet import PacketError, build_packet
        cfg = {
            "business_type": "consignment", "base_dir": str(self.base), "master_path": str(self.master),
            "party_a": {"name": "脱敏甲方"}, "party_b": {"name": "脱敏乙方"},
            "contract": {"no": "TEST-2026", "method": "铺货代销", "basis": "实销实结"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "sources": {"jia": self.jia.name, "sales": self.sales.name, "inventory": self.inventory.name},
        }
        previous = os.environ.get("RECON_CFG")
        os.environ["RECON_CFG"] = json.dumps(cfg, ensure_ascii=False)
        original_sales = self.sales.read_bytes()
        try:
            build_packet(str(self.out), str(self.master))
            wb = openpyxl.load_workbook(self.sales)
            wb["10月销售"]["C2"] = "另一家门店"
            wb.save(self.sales)
            with self.assertRaises(PacketError):
                build_packet(str(self.out), str(self.master))
            self.sales.write_bytes(original_sales)
            wb = openpyxl.load_workbook(self.master)
            wb.create_sheet("出账后新增")
            wb.save(self.master)
            with self.assertRaises(PacketError):
                build_packet(str(self.out), str(self.master))
        finally:
            if previous is None:
                os.environ.pop("RECON_CFG", None)
            else:
                os.environ["RECON_CFG"] = previous

    def test_mid_month_prices_shortage_and_sidecars(self):
        result, writeback, _run = self.run_engine()
        rule = result["rule_result"]
        self.assertEqual(rule["gross_sales_amount"], "3100.00")
        self.assertEqual(rule["variance_display_amount"], "165.00")
        self.assertEqual(rule["settlement_amount"], "3100.00")
        self.assertEqual(rule["decal_amount"], "180.00")
        self.assertFalse(rule["variance_included"])
        self.assertEqual(writeback["diff_qty"], 3)
        self.assertEqual(writeback["diff_amount"], 165)
        self.assertEqual(writeback["settle_amount"], 3100)
        wb = openpyxl.load_workbook(self.out, data_only=True)
        self.assertEqual(wb["07-乙方销售流水"].max_column, 14)
        self.assertEqual(wb["07-乙方销售流水"]["L6"].value, 50)
        self.assertEqual(wb["07-乙方销售流水"]["L7"].value, 55)
        self.assertEqual(wb["07-乙方销售流水"]["M6"].value, 2000)
        self.assertEqual(wb["07-乙方销售流水"]["M7"].value, 1100)
        self.assertEqual(result["workbook"]["sha256"], __import__("hashlib").sha256(self.out.read_bytes()).hexdigest())
        check = subprocess.run([sys.executable, str(SCRIPTS / "writeback_ledger.py"), str(self.master),
                                "--from-json", str(self.base / "_回写.json"), "--json"],
                               text=True, capture_output=True, timeout=15)
        self.assertEqual(check.returncode, 0, check.stdout + "\n" + check.stderr)

    def test_return_uses_original_sale_and_decal(self):
        rows = self.rows + [["R1", "2026-10-25", "一店", "脱敏商品", BAR, "10ml", "件", 5, "测试",
                            "退货", "S1", None, "是", None]]
        result, writeback, _run = self.run_engine(sales_rows=rows, inventory_qty=42)
        rule = result["rule_result"]
        self.assertEqual(rule["return_amount"], "-250.00")
        self.assertEqual(rule["sales_settlement_amount"], "2850.00")
        # 退货只登记「已计提」：按本期原销售实际计提的期末标准 3 冲减（5×3=15），
        # 不按原销售日期标准 2 猜算，期中调标不留贴花残值。
        self.assertEqual(rule["decal_return_amount"], "-15.00")
        self.assertEqual(rule["decal_amount"], "165.00")
        self.assertEqual(rule["quantity_reconciliation"][0]["net_sales_qty"], "55")
        self.assertEqual(writeback["settle_amount"], 2850)
        wb = openpyxl.load_workbook(self.out, data_only=True)
        self.assertEqual(wb["07-乙方销售流水"]["H8"].value, -5)
        self.assertEqual(wb["07-乙方销售流水"]["L8"].value, 50)
        self.assertEqual(wb["07-乙方销售流水"]["M8"].value, -250)

    def test_review_packet_rejects_stale_hash_and_refreshes_after_v17(self):
        result, _writeback, _run = self.run_engine()
        sys.path.insert(0, str(SCRIPTS))
        from apply_v17 import refresh_settlement_result
        from make_review_packet import PacketError, build_packet
        cfg = {
            "business_type": "consignment", "base_dir": str(self.base), "master_path": str(self.master),
            "party_a": {"name": "脱敏甲方"}, "party_b": {"name": "脱敏乙方"},
            "contract": {"no": "TEST-2026", "method": "铺货代销", "basis": "实销实结"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "sources": {"jia": self.jia.name, "sales": self.sales.name, "inventory": self.inventory.name},
        }
        previous = os.environ.get("RECON_CFG")
        os.environ["RECON_CFG"] = json.dumps(cfg, ensure_ascii=False)
        try:
            packet = build_packet(str(self.out), str(self.master))
            self.assertEqual(packet["business_type"], "consignment")
            self.assertEqual(packet["figures"]["opening_qty"], 100)
            self.assertEqual(packet["figures"]["settled_amount"], 3100)
            self.assertEqual(packet["figures"]["diff_qty"], 3)
            self.assertNotIn("closing_inventory_value", packet["figures"])
            self.assertEqual(len(packet["priced_rows"]), 2)
            self.assertEqual(packet["rule_result"], result["rule_result"])
            wb = openpyxl.load_workbook(self.out)
            wb["00-对账说明"]["A1"] = str(wb["00-对账说明"]["A1"].value) + "（附表已保存）"
            wb.save(self.out)
            with self.assertRaises(PacketError):
                build_packet(str(self.out), str(self.master))
            refresh_settlement_result(str(self.out))
            refreshed = build_packet(str(self.out), str(self.master))
            self.assertEqual(refreshed["workbook"]["sha256"], __import__("hashlib").sha256(self.out.read_bytes()).hexdigest())
            self.assertEqual(refreshed["rule_result"], result["rule_result"])
        finally:
            if previous is None:
                os.environ.pop("RECON_CFG", None)
            else:
                os.environ["RECON_CFG"] = previous

    def test_multiple_price_columns_are_blocked(self):
        wb = openpyxl.load_workbook(self.master)
        ws = wb["价格"]
        ws.cell(1, 6).value = "供货价"
        for row in range(2, ws.max_row + 1):
            ws.cell(row, 6).value = 999
        wb.save(self.master)
        _result, _writeback, run = self.run_engine(expect=1)
        self.assertIn("AMBIGUOUS_RATE_COLUMN", run.stderr)
        self.assertIn("不得任选一列", run.stderr)

    def test_sales_rows_missing_serial_or_name_are_blocked_with_row_number(self):
        for column, label in ((0, "流水号"), (3, "商品名称")):
            with self.subTest(missing=label):
                rows = [list(row) for row in self.rows]
                rows[1][column] = None
                _result, _writeback, run = self.run_engine(sales_rows=rows, expect=1)
                self.assertIn("第 2 行", run.stderr)
                self.assertIn(label, run.stderr)
                self.assertIn("不得静默跳过", run.stderr)

    def test_sales_aggregate_rows_are_exempted_with_warning(self):
        rows = [list(row) for row in self.rows]
        aggregate = [None] * len(self.rows[0])
        aggregate[3], aggregate[7] = "合计", 60
        result, _writeback, run = self.run_engine(sales_rows=rows + [aggregate])
        self.assertEqual(result["rule_result"]["quantity_reconciliation"][0]["net_sales_qty"], "60")
        self.assertIn("豁免 1 行", run.stdout + run.stderr)

    def test_non_settlement_bip_and_inventory_do_not_need_price_or_create_variance(self):
        other = "690000000002"
        wb = openpyxl.load_workbook(self.master)
        wb["SKU"].append([other, "非结算赠品", "非结算赠品", "", "件", "测试", "非结算类"])
        wb.save(self.master)
        wb = openpyxl.load_workbook(self.jia)
        wb["BIP挂账"].append(["非结算赠品", 500, 500, "正装", ""])
        wb.save(self.jia)
        wb = openpyxl.load_workbook(self.inventory)
        wb["库存"].append(["非结算赠品", other, "", "测试", 999, 0, "测试"])
        wb.save(self.inventory)
        result, writeback, _run = self.run_engine()
        self.assertEqual(len(result["rule_result"]["quantity_reconciliation"]), 1)
        self.assertEqual(result["rule_result"]["quantity_reconciliation"][0]["barcode"], BAR)
        self.assertEqual(writeback["opening_qty"], 100)
        self.assertEqual(writeback["diff_qty"], 3)
        self.assertEqual(writeback["diff_amount"], 165)
        self.assertEqual(writeback["settle_amount"], 3100)

    def test_statement_no_and_contract_provenance_are_preserved(self):
        result, writeback, _run = self.run_engine(statement_no="DJ-202610-01")
        self.assertEqual(writeback["contract_no"], "TEST-2026")
        self.assertEqual(writeback["business_type"], "铺货代销")
        self.assertEqual(writeback["statement_no"], "DJ-202610-01")
        self.assertEqual(result["writeback"], writeback)
        wb = openpyxl.load_workbook(self.out, data_only=True)
        self.assertIn("DJ-202610-01", wb["01-对账单"]["E5"].value)

    def test_review_packet_prefers_result_credential_source_paths(self):
        # P2-13：cfg.sources 文件名与结果凭证不一致时，复核包采用凭证中记录的实际绝对路径，
        # 不得按 cfg.base_dir + cfg.sources 重算路径而误拒。
        self.run_engine()
        sys.path.insert(0, str(SCRIPTS))
        from make_review_packet import build_packet
        cfg = {
            "business_type": "consignment", "base_dir": "/nonexistent-dir",
            "master_path": str(self.master),
            "party_a": {"name": "脱敏甲方"}, "party_b": {"name": "脱敏乙方"},
            "contract": {"no": "TEST-2026", "method": "铺货代销", "basis": "实销实结"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "sources": {"jia": "不存在的文件.xlsx", "sales": "另一个.xlsx", "inventory": "还有一个.xlsx"},
        }
        previous = os.environ.get("RECON_CFG")
        os.environ["RECON_CFG"] = json.dumps(cfg, ensure_ascii=False)
        try:
            packet = build_packet(str(self.out), str(self.master))
            self.assertEqual(packet["figures"]["diff_qty"], 3)
            for source in packet["sources"]:
                self.assertTrue(source["sha256"], source)
        finally:
            if previous is None:
                os.environ.pop("RECON_CFG", None)
            else:
                os.environ["RECON_CFG"] = previous

    def test_review_packet_cli_accepts_cfg_without_environment(self):
        cfg_path = self.base / "RECON_CFG.json"
        self.run_engine(cfg_out=cfg_path)
        packet_path = self.base / "packet.json"
        md_path = self.base / "packet.md"
        env = dict(os.environ)
        env.pop("RECON_CFG", None)
        run = subprocess.run([sys.executable, str(SCRIPTS / "make_review_packet.py"), str(self.out),
                              "--cfg", str(cfg_path), "--master", str(self.master),
                              "--out", str(packet_path), "--md", str(md_path)],
                             env=env, cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        packet = json.loads(packet_path.read_text(encoding="utf-8"))
        self.assertEqual(packet["contract"]["no"], "TEST-2026")
        self.assertEqual(packet["period"]["start"], "2026-10-01")
        self.assertTrue(all(row["sha256"] for row in packet["sources"]))
        self.assertTrue(md_path.exists())

    def test_variance_authorization_applies_to_this_period_only(self):
        auth = {"approved": True, "period_start": "2026-10-01", "period_end": "2026-10-31"}
        result, writeback, _run = self.run_engine(authorization=auth)
        self.assertTrue(result["rule_result"]["variance_included"])
        self.assertEqual(result["rule_result"]["settlement_amount"], "3265.00")
        self.assertEqual(writeback["settle_amount"], 3265)
        wb = openpyxl.load_workbook(self.out, data_only=True)
        ws = wb["01-对账单"]
        total_row = next(cell.row for cell in ws["A"] if cell.value == "本期结算总额")
        self.assertEqual(ws.cell(total_row, 7).value, 3265)
        self.assertEqual(ws.cell(total_row + 1, 4).value, "叁仟贰佰陆拾伍元整")


if __name__ == "__main__":
    unittest.main()
