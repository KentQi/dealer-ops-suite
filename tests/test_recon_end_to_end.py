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
sys.path.insert(0, str(SCRIPTS))

from koujing import blank as koujing_blank, digest as koujing_digest  # noqa: E402

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
        self.jia = self.base / "甲方期初.xlsx"
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
            "甲方挂账": [["商品名称", "未结算数量", "初始数量", "物料属性", "物料编码"],
                        ["脱敏商品", 100, 100, "正装", ""]],
        })
        workbook(self.inventory, {
            "库存": [["商品名称", "条码", "规格", "分类", "库存数量", "销售价", "品牌"],
                     ["脱敏商品", BAR, "10ml", "测试", 37, 100, "测试"]],
        })
        # 口径夹具：与真实 load_master 强制问卡同构，铺货出账凭证必须携带口径绑定。
        record = koujing_blank(scene="每月出账", period="2026-10", asof="2026-10-31",
                               customer="脱敏测试")
        record["confirmed_by"] = "测试"
        record["decisions"]["usage"] = {
            "values": {"试用装": 0},
            "ledger": [{"name": "试用装", "kind": "试用装", "opening": 0, "shipped": 0, "used": 0}]}
        record["internal_only"] = ["内部结算口径备注"]
        record["config_sha256"] = koujing_digest(record)
        self.koujing_record = record
        self.koujing_path = self.base / "_口径确认.json"
        self.koujing_path.write_text(json.dumps(record, ensure_ascii=False, indent=2) + "\n",
                                     encoding="utf-8")
        self.rows = [
            ["流水号", "销售时间", "门店", "商品名称", "条码", "规格", "单位", "销售数量", "品牌",
             "业务类型", "原销售流水号", "原销售单价", "原销售已计提贴花", "原销售贴花标准"],
            ["S1", "2026-10-10", "一店", "脱敏商品", BAR, "10ml", "件", 40, "测试", "销售", None, None, None, None],
            ["S2", "2026-10-20", "一店", "脱敏商品", BAR, "10ml", "件", 20, "测试", "销售", None, None, None, None],
        ]

    def tearDown(self):
        self._tmp.cleanup()

    def run_engine(self, *, sales_rows=None, inventory_qty=37, authorization=None,
                   statement_no=None, cfg_out=None, expect=0, v17=False):
        workbook(self.sales, {"10月销售": sales_rows or self.rows})
        wb = openpyxl.load_workbook(self.inventory)
        wb["库存"]["E2"] = inventory_qty
        wb.save(self.inventory)
        cfg = {
            "business_type": "consignment",
            "koujing": self.koujing_record,
            "party_a": {"name": "脱敏甲方", "tax": "TEST-A"},
            "party_b": {"name": "脱敏乙方", "tax": "TEST-B"},
            "bank": {"account_name": "脱敏甲方", "account_no": "000001", "bank_name": "测试银行"},
            "contract": {"no": "TEST-2026", "method": "铺货代销", "basis": "实销实结"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "base_dir": str(self.base), "master_path": str(self.master),
            "sources": {"jia": self.jia.name, "sales": self.sales.name, "inventory": self.inventory.name},
            "out_file": str(self.out), "overwrite": True,
            "v17": {"enabled": bool(v17), "koujing_path": str(self.koujing_path)} if v17
            else {"enabled": False},
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
        cases = ((self.jia, "甲方挂账", "B2", {}),
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
            "koujing": self.koujing_record,
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

    def test_review_packet_rejects_koujing_changed_after_settlement(self):
        # 口径绑定（复核链路 koujing_sha256）：出账后改 _口径确认.json，复核包生成必须拒绝；
        # 正常路径下凭证与复核包都要携带口径绑定。
        result, _writeback, _run = self.run_engine()
        self.assertTrue(result["koujing"]["sha256"], "铺货凭证必须携带口径哈希")
        self.assertEqual(result["koujing"]["path"], str(self.koujing_path))
        original = self.koujing_path.read_bytes()
        sys.path.insert(0, str(SCRIPTS))
        from make_review_packet import PacketError, build_packet
        cfg = {
            "business_type": "consignment", "base_dir": str(self.base), "master_path": str(self.master),
            "koujing": self.koujing_record,
            "party_a": {"name": "脱敏甲方"}, "party_b": {"name": "脱敏乙方"},
            "contract": {"no": "TEST-2026", "method": "铺货代销", "basis": "实销实结"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "sources": {"jia": self.jia.name, "sales": self.sales.name, "inventory": self.inventory.name},
        }
        previous = os.environ.get("RECON_CFG")
        os.environ["RECON_CFG"] = json.dumps(cfg, ensure_ascii=False)
        try:
            packet = build_packet(str(self.out), str(self.master))
            self.assertEqual(packet["koujing"]["sha256"], result["koujing"]["sha256"])
            # 出账后改口径：任一内容字段变化都会改变 koujing.digest
            tampered = json.loads(original.decode("utf-8"))
            tampered["decisions"]["settlement_basis"]["field"] = "初始数量"
            self.koujing_path.write_text(json.dumps(tampered, ensure_ascii=False, indent=2) + "\n",
                                         encoding="utf-8")
            with self.assertRaises(PacketError) as ctx:
                build_packet(str(self.out), str(self.master))
            self.assertIn("口径", str(ctx.exception))
        finally:
            self.koujing_path.write_bytes(original)
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
            "koujing": self.koujing_record,
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
            rule_ref = packet["rule_result"]
            self.assertEqual(set(rule_ref), {"_external", "sha256", "rows", "business_type"})
            self.assertEqual(rule_ref["_external"], "202610铺货代销对账_复核包_rule_result.json")
            self.assertEqual(rule_ref["rows"], 2)
            self.assertEqual(rule_ref["business_type"], "consignment")
            sidecar = self.base / "_复核记录" / rule_ref["_external"]
            with open(sidecar, encoding="utf-8") as fh:
                self.assertEqual(json.load(fh), result["rule_result"])
            self.assertEqual(rule_ref["sha256"],
                             __import__("hashlib").sha256(sidecar.read_bytes()).hexdigest())
            # issue #6 防回填钉死：rule_result 逐笔内容不得回流包内，体积收敛不被悄悄撤销
            packet_text = json.dumps(packet)
            self.assertNotIn("quantity_reconciliation", packet_text)
            self.assertNotIn("priced_lines", packet_text)
            wb = openpyxl.load_workbook(self.out)
            wb["00-对账说明"]["A1"] = str(wb["00-对账说明"]["A1"].value) + "（附表已保存）"
            wb.save(self.out)
            with self.assertRaises(PacketError):
                build_packet(str(self.out), str(self.master))
            refresh_settlement_result(str(self.out))
            refreshed = build_packet(str(self.out), str(self.master))
            self.assertEqual(refreshed["workbook"]["sha256"], __import__("hashlib").sha256(self.out.read_bytes()).hexdigest())
            # v17 只同步工作簿路径与哈希，rule_result 不变 → 侧车内容仍等于凭证内规则底稿
            refreshed_ref = refreshed["rule_result"]
            self.assertEqual(refreshed_ref["_external"], rule_ref["_external"])
            self.assertEqual(refreshed_ref["sha256"], rule_ref["sha256"])
            with open(self.base / "_复核记录" / refreshed_ref["_external"], encoding="utf-8") as fh:
                self.assertEqual(json.load(fh), result["rule_result"])
        finally:
            if previous is None:
                os.environ.pop("RECON_CFG", None)
            else:
                os.environ["RECON_CFG"] = previous

    def _packet_with_credential(self, out_path):
        """出账后按默认配置生成新格式复核包（rule_result 外置），返回 packet dict。"""
        sys.path.insert(0, str(SCRIPTS))
        from make_review_packet import build_packet
        cfg = {
            "business_type": "consignment", "base_dir": str(self.base), "master_path": str(self.master),
            "koujing": self.koujing_record,
            "party_a": {"name": "脱敏甲方"}, "party_b": {"name": "脱敏乙方"},
            "contract": {"no": "TEST-2026", "method": "铺货代销", "basis": "实销实结"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "sources": {"jia": self.jia.name, "sales": self.sales.name, "inventory": self.inventory.name},
        }
        previous = os.environ.get("RECON_CFG")
        os.environ["RECON_CFG"] = json.dumps(cfg, ensure_ascii=False)
        try:
            packet = build_packet(str(self.out), str(self.master), out_path=str(out_path))
        finally:
            if previous is None:
                os.environ.pop("RECON_CFG", None)
            else:
                os.environ["RECON_CFG"] = previous
        out_path.write_text(json.dumps(packet, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return packet

    def _pass_report(self, packet, path):
        """按复核包生成一份 R1–R10 全通过的合法报告。"""
        report = {
            "schema": "dealer-ops/review-report@1",
            "reviewer": {"name": "测试复核", "model": "test"},
            "reviewed_at": "2026-11-20T08:00:00+08:00",
            "workbook_sha256": packet["workbook"]["sha256"],
            "verdict": "通过",
            "checklist": [{"id": "R%d" % index, "result": "通过", "note": "已核对"}
                          for index in range(1, 11)],
            "findings": [], "summary": "复核通过",
            "period": {"start": "2026-10-01", "end": "2026-10-31"},
        }
        path.write_text(json.dumps(report, ensure_ascii=False), encoding="utf-8")
        return report

    def _run_apply(self, report_path, packet_path):
        return subprocess.run([sys.executable, str(SCRIPTS / "apply_review.py"), str(self.out),
                               str(report_path), "--packet", str(packet_path), "--apply"],
                              text=True, capture_output=True, timeout=30)

    def test_apply_review_rejects_missing_or_tampered_rule_result_sidecar(self):
        result, _writeback, _run = self.run_engine()
        packet_path = self.base / "packet.json"
        report_path = self.base / "report.json"
        packet = self._packet_with_credential(packet_path)
        sidecar = self.base / packet["rule_result"]["_external"]
        self.assertTrue(sidecar.exists())
        self._pass_report(packet, report_path)

        sidecar.unlink()  # 侧车缺失：须拒绝且不写复核表/复核锁
        run = self._run_apply(report_path, packet_path)
        self.assertNotEqual(run.returncode, 0, run.stdout + run.stderr)
        self.assertIn("复核包引用的规则底稿不存在", run.stderr)
        self.assertIn("packet_rule_result.json", run.stderr)
        self.assertFalse((self.base / "_复核记录" / "复核锁.json").exists())

        sidecar.write_text(json.dumps(result["rule_result"], ensure_ascii=False) + "\n",
                           encoding="utf-8")  # 字节变了（非缩进重写）→ 哈希不符
        run = self._run_apply(report_path, packet_path)
        self.assertNotEqual(run.returncode, 0, run.stdout + run.stderr)
        self.assertIn("规则底稿与复核包不一致", run.stderr)
        self.assertIn("重新生成复核包", run.stderr)
        self.assertFalse((self.base / "_复核记录" / "复核锁.json").exists())

    def test_legacy_inline_rule_result_packet_still_applies(self):
        # 旧格式（rule_result 整体内联的历史归档包）不做侧车校验，仍可 --apply。
        # 真实流程对齐（T013）：复核发生在 apply_v17 渲染与门禁之后——原始引擎簿含
        # 「请在挂账表补…」类内部文案，--apply 终态复扫会正确拦截；先渲染定稿再验兼容。
        result, _writeback, _run = self.run_engine(v17=True)
        cfg_path = self.base / "cfg_v17.json"
        cfg_path.write_text(json.dumps({
            "business_type": "consignment",
            "koujing": self.koujing_record,
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "base_dir": str(self.base), "master_path": str(self.master),
            "out_file": str(self.out), "overwrite": True,
            "v17": {"enabled": True, "mode": "full"},
        }, ensure_ascii=False), encoding="utf-8")
        v17_run = subprocess.run([sys.executable, str(SCRIPTS / "apply_v17.py"),
                                  "--cfg", str(cfg_path)],
                                 cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(v17_run.returncode, 0, v17_run.stdout + v17_run.stderr)
        report_path = self.base / "report.json"
        packet = self._packet_with_credential(self.base / "legacy_packet.json")
        legacy = dict(packet)
        legacy["rule_result"] = result["rule_result"]
        legacy_path = self.base / "legacy_packet.json"
        legacy_path.write_text(json.dumps(legacy, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        self._pass_report(packet, report_path)
        run = self._run_apply(report_path, legacy_path)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        self.assertIn("[完成] 复核有效", run.stdout)

    def test_no_credential_mode_writes_no_sidecar(self):
        self.run_engine()
        legacy_dir = self.base / "无凭证"
        legacy_dir.mkdir()
        legacy_workbook = legacy_dir / "2025历史工作簿.xlsx"
        legacy_workbook.write_bytes(self.out.read_bytes())
        sys.path.insert(0, str(SCRIPTS))
        from make_review_packet import build_packet
        packet = build_packet(str(legacy_workbook))
        self.assertIsNone(packet["rule_result"])
        self.assertEqual(packet["priced_rows"], [])
        self.assertTrue(any("未找到同目录 _结算结果.json" in warning for warning in packet["warnings"]))
        self.assertFalse((legacy_dir / "_复核记录" / "2025历史工作簿_复核包_rule_result.json").exists())

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

    def test_non_settlement_opening_and_inventory_do_not_need_price_or_create_variance(self):
        other = "690000000002"
        wb = openpyxl.load_workbook(self.master)
        wb["SKU"].append([other, "非结算赠品", "非结算赠品", "", "件", "测试", "非结算类"])
        wb.save(self.master)
        wb = openpyxl.load_workbook(self.jia)
        wb["甲方挂账"].append(["非结算赠品", 500, 500, "正装", ""])
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
            "koujing": self.koujing_record,
            "v17": {"enabled": False, "koujing_path": str(self.koujing_path)},
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
