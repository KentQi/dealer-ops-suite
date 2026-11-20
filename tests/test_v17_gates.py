#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v1.7 交付门禁钉子：效期开关（09 表）与增量冻结（01–07 逐格不变）。

夹具与 tests/test_recon_end_to_end.py 的 ConsignmentWorkbookTest 同构：
主档三表 + 甲方期初/乙方销售/乙方库存三导出 + koujing 口径夹具，
真实出账（recon_engine）后再走 apply_v17 附表渲染。
说明：引擎 01-对账单 说明 6 自带 l1 禁词「主档」（早于禁词门禁的历史文案），
门禁全量扫描目前必然失败，故本文件的渲染流程用 --no-gate 出簿，
效期硬断言（references/对外文案治理.md:89 清单）由本文件逐字符串自证。
"""

import copy
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from check_public_copy import scan as public_scan  # noqa: E402
from koujing import blank as koujing_blank, digest as koujing_digest  # noqa: E402

BAR = "690000000001"
# references/对外文案治理.md「关闭时的硬断言」：全簿命中数 = 0、无 09 表名、无悬空附表 09
DOC89_TERMS = ("效期", "批次", "到期", "临期", "预警", "剩余月数", "FIFO")
FROZEN_NOS = ("01", "02", "03", "04", "05", "06", "07")
FULL_SHEETS = ["00-对账说明", "01-对账单", "02-库存差异表", "03-源数据对照",
               "04-甲方挂账明细", "05-乙方销售汇总", "06-乙方库存明细",
               "07-乙方销售流水", "08-试用装与中小样"]


def workbook(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)


class ConsignmentFixture(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(prefix=".test-v17-gates-", dir=ROOT)
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
        self.koujing_path = self.base / "_口径确认.json"
        self.write_koujing()
        self.rows = [
            ["流水号", "销售时间", "门店", "商品名称", "条码", "规格", "单位", "销售数量", "品牌",
             "业务类型", "原销售流水号", "原销售单价", "原销售已计提贴花", "原销售贴花标准"],
            ["S1", "2026-10-10", "一店", "脱敏商品", BAR, "10ml", "件", 40, "测试", "销售", None, None, None, None],
            ["S2", "2026-10-20", "一店", "脱敏商品", BAR, "10ml", "件", 20, "测试", "销售", None, None, None, None],
        ]
        workbook(self.sales, {"10月销售": self.rows})

    def tearDown(self):
        self._tmp.cleanup()

    def write_koujing(self):
        self.koujing_path.write_text(
            json.dumps(self.koujing_record, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8")

    def set_expiry(self, enabled):
        """效期口径开关变更后必须重算口径哈希（口径门禁的绑定要求）。"""
        self.koujing_record["decisions"]["expiry"]["enabled"] = enabled
        self.koujing_record["config_sha256"] = koujing_digest(self.koujing_record)
        self.write_koujing()

    def build_cfg(self, *, mode="full", input_workbook=None):
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
            "out_file": str(self.out), "overwrite": True,
            "v17": {"enabled": True, "mode": mode},
        }
        if input_workbook is not None:
            cfg["v17"]["input_workbook"] = input_workbook
        return cfg

    def run_engine(self, *, cfg=None, expect=0):
        cfg = cfg or self.build_cfg()
        env = dict(os.environ, RECON_CFG=json.dumps(cfg, ensure_ascii=False))
        run = subprocess.run([sys.executable, str(SCRIPTS / "recon_engine.py")], env=env,
                             cwd=ROOT, text=True, capture_output=True, timeout=30)
        self.assertEqual(run.returncode, expect, run.stdout + "\n" + run.stderr)
        return run, cfg

    def dump_cfg(self, cfg, name="cfg_v17.json"):
        path = self.base / name
        path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        return str(path)

    def run_v17(self, cfg_path, *extra):
        return subprocess.run([sys.executable, str(SCRIPTS / "apply_v17.py"),
                               "--cfg", cfg_path, *extra],
                              cwd=ROOT, text=True, capture_output=True, timeout=30)

    def book_strings(self, path):
        """单元格、页眉页脚、表名与 docProps 的全量字符串（与门禁扫描同口径）。"""
        wb = openpyxl.load_workbook(path, data_only=False)
        out = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        out.append(("%s!%s" % (ws.title, cell.coordinate), cell.value))
            out.append((ws.title + "!表名", ws.title))
            for part in ("oddHeader", "oddFooter", "evenHeader", "evenFooter",
                         "firstHeader", "firstFooter"):
                obj = getattr(ws, part, None)
                if obj is None:
                    continue
                for side in ("left", "center", "right"):
                    txt = getattr(getattr(obj, side, None), "text", None)
                    if txt:
                        out.append(("%s!%s.%s" % (ws.title, part, side), txt))
        wb.close()
        with zipfile.ZipFile(path) as zf:
            for member in ("docProps/core.xml", "docProps/app.xml"):
                if member in zf.namelist():
                    out.append((member, zf.read(member).decode("utf-8", "ignore")))
        return out

    def cell_values(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        data = {name: {cell.coordinate: cell.value
                       for row in wb[name].iter_rows() for cell in row}
                for name in wb.sheetnames}
        wb.close()
        return data


class V17GateTest(ConsignmentFixture):
    def test_expiry_off_full_mode_leaves_zero_expiry_terms_and_no_sheet09(self):
        _run, cfg = self.run_engine()
        self.assertTrue((self.base / "_v17_data.json").exists(), "引擎必须产出 v17 交接数据")
        payload_path = self.dump_cfg(cfg)
        run = self.run_v17(payload_path, "--no-gate")
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        payload = json.loads(run.stdout)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["mode"], "full")
        # --no-gate 模式的回执不含 expiry_enabled；开关语义由下方簿内容断言兜底

        wb = openpyxl.load_workbook(self.out)
        self.assertEqual(wb.sheetnames, FULL_SHEETS)
        wb.close()
        # 硬断言清单（对外文案治理.md:89）：全簿效期字样命中数 = 0
        hits = [(loc, text[:60]) for loc, text in self.book_strings(self.out)
                if any(term in text for term in DOC89_TERMS)]
        self.assertEqual(hits, [])
        # 无悬空「附表 09」交叉引用
        dangling = [(loc, text[:60]) for loc, text in self.book_strings(self.out)
                    if "附表 09" in text]
        self.assertEqual(dangling, [])
        # 旧命名附表必须被替换为口径 sheet_map 命名
        self.assertNotIn("08-效期预警", FULL_SHEETS)
        self.assertNotIn("09-试用装与配赠", FULL_SHEETS)
        # 结算凭证随附表渲染同步哈希，且结算数字不受附表影响
        with open(self.base / "_结算结果.json", encoding="utf-8") as fh:
            result = json.load(fh)
        self.assertEqual(result["workbook"]["sha256"],
                         hashlib.sha256(self.out.read_bytes()).hexdigest())
        self.assertEqual(result["rule_result"]["business_type"], "consignment")
        self.assertEqual(result["rule_result"]["sales_settlement_amount"], "3100.00")
        self.assertEqual(result["rule_result"]["quantity_reconciliation"][0]["net_sales_qty"], "60")

    def test_expiry_on_full_mode_builds_sheet09(self):
        self.set_expiry(True)
        _run, cfg = self.run_engine()
        run = self.run_v17(self.dump_cfg(cfg), "--no-gate")
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        payload = json.loads(run.stdout)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["mode"], "full")
        wb = openpyxl.load_workbook(self.out)
        self.assertIn("09-效期预警", wb.sheetnames)
        self.assertIn("08-试用装与中小样", wb.sheetnames)
        wb.close()
        # 开关真的驱动了内容：簿内必须出现效期字样，且附表交叉引用可解析
        self.assertTrue(any("效期" in text for _loc, text in self.book_strings(self.out)))
        gate = public_scan(str(self.out), cfg=cfg, strict=False)
        self.assertEqual([p for p in gate["problems"] if p.startswith(("A3", "A6"))], [],
                         gate["problems"])

    def test_incremental_mode_freezes_sheets_01_to_07_cell_by_cell(self):
        _run, cfg = self.run_engine()
        run = self.run_v17(self.dump_cfg(cfg), "--no-gate")
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        snapshot = self.base / "增量前快照.xlsx"
        snapshot.write_bytes(self.out.read_bytes())
        with open(self.base / "_结算结果.json", encoding="utf-8") as fh:
            rule_before = json.load(fh)["rule_result"]

        wb = openpyxl.load_workbook(self.out)
        wb["00-对账说明"]["A1"] = "无关改动"  # 00 表不属于冻结范围，允许被覆盖重生成
        wb.save(self.out)
        wb.close()

        cfg_inc = copy.deepcopy(cfg)
        cfg_inc["v17"] = {"enabled": True, "mode": "incremental",
                          "input_workbook": str(self.out)}
        run = self.run_v17(self.dump_cfg(cfg_inc, "cfg_incremental.json"), "--no-gate")
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        payload = json.loads(run.stdout)
        self.assertEqual(payload["mode"], "incremental")

        before = self.cell_values(snapshot)
        after = self.cell_values(self.out)
        self.assertEqual(list(before), list(after), "增量不得增删或重排工作表")
        diffs = []
        for no in FROZEN_NOS:
            name = next(n for n in before if n.startswith(no + "-"))
            for coordinate, value in before[name].items():
                if after[name].get(coordinate) != value:
                    diffs.append((name, coordinate, value, after[name].get(coordinate)))
        self.assertEqual(diffs[:5], [])
        self.assertEqual(diffs, [])
        self.assertTrue(any("_增量前_" in p.name for p in (self.base / "_备份").iterdir()),
                        "增量前必须有备份")
        with open(self.base / "_结算结果.json", encoding="utf-8") as fh:
            result = json.load(fh)
        self.assertEqual(result["rule_result"], rule_before)
        self.assertEqual(result["workbook"]["sha256"],
                         hashlib.sha256(self.out.read_bytes()).hexdigest())


class PipelineGateTest(ConsignmentFixture):
    """对外文案门禁是出账流水线的强制环节（T011）。

    历史教训：01-对账单说明#6 曾自带 l1 禁词「主档」（T009 修复，d234f6a），
    正常出账过不了闸，全流程被迫 --no-gate 绕行。本类把「闸门恢复实职」
    钉成回归：不带 --no-gate 必须实际扫描并放行干净出账；禁词注入必须拒绝。
    """

    def test_clean_statement_passes_gate_without_no_gate_flag(self):
        _run, cfg = self.run_engine()
        run = self.run_v17(self.dump_cfg(cfg))  # 不带 --no-gate
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        payload = json.loads(run.stdout)
        self.assertTrue(payload["ok"])
        self.assertGreater(payload["gate"], 0, "门禁必须实际扫描（回执含扫描计数）")

    def test_gate_rejects_banned_word_injected_after_settlement(self):
        _run, cfg = self.run_engine()
        # 出账后、交付渲染前，向对账单注入 l1 禁词（模拟手工改动或上游污染）
        wb = openpyxl.load_workbook(self.out)
        wb["01-对账单"]["H60"] = "补充说明：本表数量取自挂账表"
        wb.save(self.out)
        wb.close()
        run = self.run_v17(self.dump_cfg(cfg))
        self.assertNotEqual(run.returncode, 0, run.stdout + run.stderr)
        self.assertIn("对外门禁未通过", run.stderr)


if __name__ == "__main__":
    unittest.main()
