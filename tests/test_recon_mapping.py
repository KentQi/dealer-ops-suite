#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""铺货侧结构映射链路钉子：乙导出改列名/表名 → inspect_schema → apply_mapping → 带映射重跑。

场景（原 examples/schema_demo.py 的断言收编）：
  1) 标准命名出账 → 基线金额；
  2) 「乙方库存.xlsx」改名「库存导出.xlsx」（表名与「条码/库存数量」列名同时漂移），
     cfg sources 指新文件且无映射 → 引擎阻断、不落盘；
  3) inspect_schema 出映射建议 → apply_mapping --write 落 _数据映射.json（含确认记录归档）
     → 引擎带映射重跑 → 结算金额与基线一致。
映射注入方式与 load_master.py 一致：发现 _数据映射.json 后读入并放进 RECON_CFG["mapping"]。
"""

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
MAPPING_SCHEMA = "dealer-ops/data-mapping@1"


def workbook(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


class ConsignmentFixture(unittest.TestCase):
    """与 tests/test_recon_end_to_end.py 的 ConsignmentWorkbookTest 同构的出账夹具。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(prefix=".test-mapping-", dir=ROOT)
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
        record = koujing_blank(scene="每月出账", period="2026-10", asof="2026-10-31",
                               customer="脱敏测试")
        record["confirmed_by"] = "测试"
        record["decisions"]["usage"] = {
            "values": {"试用装": 0},
            "ledger": [{"name": "试用装", "kind": "试用装", "opening": 0, "shipped": 0, "used": 0}]}
        record["internal_only"] = ["内部结算口径备注"]
        record["config_sha256"] = koujing_digest(record)
        self.koujing_record = record
        self.rows = [
            ["流水号", "销售时间", "门店", "商品名称", "条码", "规格", "单位", "销售数量", "品牌",
             "业务类型", "原销售流水号", "原销售单价", "原销售已计提贴花", "原销售贴花标准"],
            ["S1", "2026-10-10", "一店", "脱敏商品", BAR, "10ml", "件", 40, "测试", "销售", None, None, None, None],
            ["S2", "2026-10-20", "一店", "脱敏商品", BAR, "10ml", "件", 20, "测试", "销售", None, None, None, None],
        ]

    def tearDown(self):
        self._tmp.cleanup()

    def build_cfg(self, *, mapping=None):
        cfg = {
            "business_type": "consignment",
            "koujing": self.koujing_record,
            "party_a": {"name": "脱敏甲方", "tax": "TEST-A"},
            "party_b": {"name": "脱敏乙方", "tax": "TEST-B"},
            "bank": {"account_name": "脱敏甲方", "account_no": "000001", "bank_name": "测试银行"},
            "contract": {"no": "TEST-2026", "method": "铺货代销", "basis": "实销实结"},
            "period": {"start": "2026-10-01", "end": "2026-10-31", "asof": "2026-10-31"},
            "base_dir": str(self.base), "master_path": str(self.master),
            "sources": {"jia": self.jia.name, "sales": self.sales.name,
                        "inventory": self.inventory.name},
            "out_file": str(self.out), "overwrite": True, "v17": {"enabled": False},
        }
        if mapping is not None:
            cfg["mapping"] = mapping
        return cfg

    def run_engine(self, *, cfg=None, expect=0):
        workbook(self.sales, {"10月销售": self.rows})
        cfg = self.build_cfg() if cfg is None else cfg
        env = dict(os.environ, RECON_CFG=json.dumps(cfg, ensure_ascii=False))
        run = subprocess.run([sys.executable, str(SCRIPTS / "recon_engine.py")], env=env,
                             cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(run.returncode, expect, run.stdout + "\n" + run.stderr)
        return run

    def read_rule_and_writeback(self):
        with open(self.base / "_结算结果.json", encoding="utf-8") as fh:
            rule = json.load(fh)["rule_result"]
        with open(self.base / "_回写.json", encoding="utf-8") as fh:
            writeback = json.load(fh)
        return rule, writeback


class ReconMappingTest(ConsignmentFixture):
    def test_standard_naming_produces_baseline_figures(self):
        self.run_engine()
        rule, writeback = self.read_rule_and_writeback()
        self.assertEqual(rule["sales_settlement_amount"], "3100.00")
        self.assertEqual(rule["quantity_reconciliation"][0]["net_sales_qty"], "60")
        self.assertEqual(writeback["settle_amount"], 3100)
        self.assertEqual(writeback["settle_qty"], 60)

    def test_renamed_export_without_mapping_is_blocked(self):
        self.rename_inventory_export()
        run = self.run_engine(expect=1)  # cfg sources 已指向新文件，但没有任何映射
        self.assertIn("未找到", run.stderr)
        self.assertIn("条码", run.stderr)
        self.assertFalse(self.out.exists(), "阻断时不得生成对账工作簿")

    def rename_inventory_export(self):
        """把乙库存导出整体漂移：文件名、工作表名、条码列与库存数量列全部改名。"""
        renamed = self.base / "库存导出.xlsx"
        os.replace(self.inventory, renamed)
        wb = openpyxl.load_workbook(renamed)
        wb["库存"].title = "在库明细"
        ws = wb["在库明细"]
        ws["B1"] = "条形码"    # 引擎关键词「条码/编码」全部落空
        ws["E1"] = "结存数"    # 引擎关键词「库存/数量」全部落空
        wb.save(renamed)
        wb.close()
        self.inventory = renamed

    def test_inspect_schema_then_apply_mapping_restores_baseline_result(self):
        self.run_engine()
        baseline_rule, baseline_writeback = self.read_rule_and_writeback()
        self.assertEqual(baseline_rule["sales_settlement_amount"], "3100.00")
        self.rename_inventory_export()

        inspect_out = self.base / "探查.json"
        run = subprocess.run([sys.executable, str(SCRIPTS / "inspect_schema.py"),
                              str(self.inventory), "--role", "inventory",
                              "--out", str(inspect_out)],
                             cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        report = json.loads(inspect_out.read_text(encoding="utf-8"))
        proposed = report["proposed_mapping"]
        self.assertEqual(proposed["schema"], MAPPING_SCHEMA)
        entry = proposed["sources"]["inventory"]
        self.assertEqual(entry["sheet"], "在库明细")
        self.assertEqual(entry["header_row"], 1)
        self.assertEqual(entry["columns"]["name"], "商品名称")
        self.assertEqual(entry["columns"]["bar"], "条形码")
        self.assertEqual(entry["columns"]["qty"], "结存数")

        proposal_path = self.base / "映射建议.json"
        proposal_path.write_text(json.dumps(proposed, ensure_ascii=False, indent=2) + "\n",
                                 encoding="utf-8")
        run = subprocess.run([sys.executable, str(SCRIPTS / "apply_mapping.py"), str(self.base),
                              str(proposal_path), "--write", "--confirmed-by", "测试"],
                             cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        mapping_file = self.base / "_数据映射.json"
        self.assertTrue(mapping_file.exists())
        confirmed = json.loads(mapping_file.read_text(encoding="utf-8"))
        self.assertEqual(confirmed["schema"], MAPPING_SCHEMA)
        self.assertEqual(confirmed["confirmed_by"], "测试")
        archives = list((self.base / "_确认记录").glob("映射_*.json"))
        self.assertTrue(archives, "映射确认必须留档")

        # 引擎带映射重跑：映射注入口径与 load_master.py 相同（cfg["mapping"]）
        cfg = self.build_cfg(mapping=confirmed)
        self.run_engine(cfg=cfg)
        rule, writeback = self.read_rule_and_writeback()
        self.assertEqual(rule["sales_settlement_amount"], "3100.00")
        self.assertEqual(rule["quantity_reconciliation"][0]["net_sales_qty"], "60")
        self.assertEqual(rule, baseline_rule, "带映射结果必须与标准命名完全一致")
        self.assertEqual(writeback, baseline_writeback)


if __name__ == "__main__":
    unittest.main()
