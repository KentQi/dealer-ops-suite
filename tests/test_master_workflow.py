#!/usr/bin/env python3
# -*- coding: utf-8 -*-
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
import make_master

BAR = "690000000001"


def save(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


class MasterWorkflowTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix=".test-master-", dir=ROOT)
        self.root = Path(self.tmp.name)
        self.master = self.root / "_主档.xlsx"
        save(self.master, {
            "主体": [["角色", "名称", "开户行", "账号"],
                     ["甲方", "甲方公司", "测试银行", "123"], ["乙方", "乙方公司", "", ""]],
            "合同": [["合同编号", "业务类型", "起始", "终止", "结算方式", "付款日", "买断日", "付款截止日"],
                     ["HT-1", "铺货代销", "2026-01-01", "2026-12-31", "铺货代销", "次月10日前", "", ""]],
            "价格": [["商品条码", "商品名称", "供货价", "生效起", "生效止"],
                     [BAR, "商品", 50, "2026-01-01", "2026-10-31"]],
            "贴花": [["商品条码", "商品名称", "贴花标准", "生效起", "生效止"]],
            "SKU": [["商品条码", "标准品名"], [BAR, "商品"]],
        })

    def tearDown(self):
        self.tmp.cleanup()

    def load_cfg(self, *extra):
        run = subprocess.run([sys.executable, str(SCRIPTS / "load_master.py"), str(self.master),
                              "--period", "202610", "--allow-legacy", *extra],
                             text=True, capture_output=True, timeout=20)
        return run, json.loads(run.stdout) if run.returncode == 0 else None

    def test_missing_period_directory_never_falls_back(self):
        run, _ = self.load_cfg()
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("不会回退", run.stderr)

    def test_consignment_selects_covering_contract_and_does_not_default_payday(self):
        (self.root / "202610").mkdir()
        wb = openpyxl.load_workbook(self.master)
        wb["合同"].append(["HT-2", "铺货代销", "2025-01-01", "2025-12-31", "铺货代销", "次月20日前", "", ""])
        wb.save(self.master)
        wb = openpyxl.load_workbook(self.master)
        wb["合同"]["F2"] = None
        wb.save(self.master)
        run, _ = self.load_cfg("--strict")
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("不会默认次月 5 日", run.stderr)
        cfg_path = self.root / "RECON_CFG.json"
        run, cfg = self.load_cfg("--pay", "2026-11-15", "--cfg-out", str(cfg_path))
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertEqual(cfg["contract"]["no"], "HT-1")
        self.assertIn("15", cfg["period"]["pay_cn"])
        self.assertEqual(json.loads(cfg_path.read_text(encoding="utf-8")), cfg)

    def test_multiple_covering_contracts_block(self):
        (self.root / "202610").mkdir()
        wb = openpyxl.load_workbook(self.master)
        wb["合同"].append(["HT-2", "铺货代销", "2026-01-01", "2026-12-31", "铺货代销", "次月20日前", "", ""])
        wb.save(self.master)
        run, _ = self.load_cfg()
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("多份合同", run.stderr)

    def test_monthly_route_requires_dates_and_uses_independent_sources(self):
        (self.root / "202610").mkdir()
        wb = openpyxl.load_workbook(self.master)
        wb["合同"]["B2"] = "月结买断"
        wb["合同"]["E2"] = "月结买断"
        wb.save(self.master)
        run, _ = self.load_cfg()
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("买断日", run.stderr)
        wb = openpyxl.load_workbook(self.master)
        wb["合同"]["G2"] = "2026-11-05"
        wb["合同"]["H2"] = "2026-11-15"
        wb.save(self.master)
        run, cfg = self.load_cfg()
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertEqual(cfg["business_type"], "monthly_buyout")
        self.assertEqual(cfg["contract"]["buyout_date"], "2026-11-05")
        self.assertEqual(cfg["contract"]["payment_due_date"], "2026-11-15")
        self.assertFalse(cfg["v17"]["enabled"])
        self.assertEqual(set(cfg["sources"]), {"receipts", "returns", "sales"})

    def test_multiple_source_files_and_monthly_pay_override_block(self):
        period = self.root / "202610"
        period.mkdir()
        save(period / "BIP-A.xlsx", {"Sheet": [["值"], [1]]})
        save(period / "BIP-B.xlsx", {"Sheet": [["值"], [2]]})
        run, _ = self.load_cfg()
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("源文件歧义", run.stderr)
        (period / "BIP-B.xlsx").unlink()
        wb = openpyxl.load_workbook(self.master)
        wb["合同"]["B2"] = "月结买断"
        wb["合同"]["E2"] = "月结买断"
        wb["合同"]["G2"] = "2026-11-05"
        wb["合同"]["H2"] = "2026-11-15"
        wb.save(self.master)
        run, _ = self.load_cfg("--pay", "2026-11-20")
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("不能用 --pay", run.stderr)

    def apply(self, fields, force=False):
        candidate = self.root / "candidate.json"
        candidate.write_text(json.dumps({"fields": fields}, ensure_ascii=False), encoding="utf-8")
        cmd = [sys.executable, str(SCRIPTS / "apply_contract.py"), str(self.master),
               str(candidate), "--write"] + (["--force"] if force else [])
        return subprocess.run(cmd, text=True, capture_output=True, timeout=20)

    def test_contract_update_requires_same_contract_number(self):
        run = self.apply([{"field": "付款日", "target": "合同.付款日", "value": "次月25日前"}])
        self.assertNotEqual(run.returncode, 0)
        self.assertIn("合同编号", run.stdout)
        run = self.apply([{"field": "付款日", "target": "合同.付款日@合同编号=HT-1", "value": "次月25日前"}], force=True)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        wb = openpyxl.load_workbook(self.master, data_only=True)
        self.assertEqual(wb["合同"]["F2"].value, "次月25日前")

    def test_new_policy_row_keeps_old_version_and_is_idempotent(self):
        fields = [
            {"field": "商品名称", "target": f"价格.商品名称@商品条码={BAR}@生效起=2026-11-01", "value": "商品"},
            {"field": "供货价", "target": f"价格.供货价@商品条码={BAR}@生效起=2026-11-01", "value": 55},
        ]
        run = self.apply(fields)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        wb = openpyxl.load_workbook(self.master, data_only=True)
        self.assertEqual(wb["价格"]["C2"].value, 50)
        self.assertEqual(wb["价格"]["C3"].value, 55)
        self.assertEqual(wb["价格"]["A3"].value, BAR)
        self.assertEqual(wb["价格"]["D3"].value, "2026-11-01")
        run = self.apply(fields)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        wb = openpyxl.load_workbook(self.master, data_only=True)
        self.assertEqual(wb["价格"].max_row, 3)

    def test_new_contract_does_not_overwrite_old(self):
        run = self.apply([
            {"field": "合同编号", "target": "合同.合同编号", "value": "HT-2"},
            {"field": "业务类型", "target": "合同.业务类型@合同编号=HT-2", "value": "月结买断"},
            {"field": "起始", "target": "合同.起始@合同编号=HT-2", "value": "2026-11-01"},
            {"field": "终止", "target": "合同.终止@合同编号=HT-2", "value": "2027-10-31"},
            {"field": "买断日", "target": "合同.买断日@合同编号=HT-2", "value": "2026-12-05"},
            {"field": "付款截止日", "target": "合同.付款截止日@合同编号=HT-2", "value": "2026-12-15"},
        ])
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        wb = openpyxl.load_workbook(self.master, data_only=True)
        self.assertEqual(wb["合同"]["A2"].value, "HT-1")
        self.assertEqual(wb["合同"]["A3"].value, "HT-2")
        self.assertEqual(wb["合同"]["G3"].value, "2026-12-05")

    def test_master_has_two_layers_and_monthly_dates(self):
        output = self.root / "new-master.xlsx"
        make_master.build(str(output))
        wb = openpyxl.load_workbook(output, data_only=True)
        self.assertIn("履约条款", wb.sheetnames)
        self.assertIn("付款截止日", [cell.value for cell in wb["期间"][1]])
        self.assertIn("原文依据", [cell.value for cell in wb["履约条款"][1]])


if __name__ == "__main__":
    unittest.main()
