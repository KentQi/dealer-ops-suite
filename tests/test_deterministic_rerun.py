#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""确定性重跑钉子：同一 cfg 两次出账，工作簿逐单元格与规则底稿完全一致。

xlsx 字节级必然不同（zip 时间戳/docProps），故按任务口径做单元格级比对：
全部工作表、全部单元格值逐一相等；_结算结果.json 的 rule_result（含
business_type 与各项 figures）与 _回写.json 两次一致。
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
        self._tmp = tempfile.TemporaryDirectory(prefix=".test-rerun-", dir=ROOT)
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

    def build_cfg(self):
        return {
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

    def run_engine(self, *, sales_rows=None, inventory_qty=37, expect=0):
        workbook(self.sales, {"10月销售": sales_rows or self.rows})
        wb = openpyxl.load_workbook(self.inventory)
        wb["库存"]["E2"] = inventory_qty
        wb.save(self.inventory)
        cfg = self.build_cfg()
        env = dict(os.environ, RECON_CFG=json.dumps(cfg, ensure_ascii=False))
        run = subprocess.run([sys.executable, str(SCRIPTS / "recon_engine.py")], env=env,
                             cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(run.returncode, expect, run.stdout + "\n" + run.stderr)
        return run

    def cell_values(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        data = {name: {cell.coordinate: cell.value
                       for row in wb[name].iter_rows() for cell in row}
                for name in wb.sheetnames}
        wb.close()
        return data

    def read_json(self, name):
        with open(self.base / name, encoding="utf-8") as fh:
            return json.load(fh)


class DeterministicRerunTest(ConsignmentFixture):
    def test_same_config_reruns_produce_identical_cells_and_rule_result(self):
        self.run_engine()
        first_book = self.base / "第一次出账.xlsx"
        first_book.write_bytes(self.out.read_bytes())
        first_rule = self.read_json("_结算结果.json")["rule_result"]
        first_writeback = self.read_json("_回写.json")

        self.run_engine()  # 同一 cfg 第二次出账，overwrite=True 覆盖同一目标

        before, after = self.cell_values(first_book), self.cell_values(self.out)
        self.assertEqual(list(before), list(after), "两次出账的工作表集合与顺序必须一致")
        diffs = []
        for name, cells in before.items():
            for coordinate, value in cells.items():
                if after[name].get(coordinate) != value:
                    diffs.append((name, coordinate, value, after[name].get(coordinate)))
        self.assertEqual(diffs[:5], [])
        self.assertEqual(diffs, [], "全部工作表单元格值必须逐格一致")

        second_rule = self.read_json("_结算结果.json")["rule_result"]
        self.assertEqual(first_rule, second_rule)
        self.assertEqual(first_rule["business_type"], "consignment")
        self.assertEqual(first_rule["sales_settlement_amount"], "3100.00")
        self.assertEqual(first_rule["quantity_reconciliation"][0]["net_sales_qty"], "60")
        self.assertEqual(first_writeback, self.read_json("_回写.json"))

    def test_changed_source_breaks_reproduction(self):
        # 对照组：比较逻辑必须真的有牙——源数据一变，逐格比对与规则底稿都要翻红
        self.run_engine()
        first_book = self.base / "基线出账.xlsx"
        first_book.write_bytes(self.out.read_bytes())
        first_rule = self.read_json("_结算结果.json")["rule_result"]

        rows = [list(row) for row in self.rows]
        rows[1][7] = 41  # S1 销量 40 → 41
        self.run_engine(sales_rows=rows)

        before, after = self.cell_values(first_book), self.cell_values(self.out)
        diffs = [(name, coordinate)
                 for name, cells in before.items()
                 for coordinate, value in cells.items()
                 if after[name].get(coordinate) != value]
        self.assertTrue(diffs, "源数据变化后工作簿必须出现单元格差异")
        second_rule = self.read_json("_结算结果.json")["rule_result"]
        self.assertNotEqual(first_rule, second_rule)
        self.assertEqual(second_rule["sales_settlement_amount"], "3150.00")


if __name__ == "__main__":
    unittest.main()
