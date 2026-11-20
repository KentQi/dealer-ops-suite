#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""对外文案门禁钉子（check_public_copy.py）。

被测 API：scan(path, cfg=None, baseline=None, strict=True) —— problems / ok 结构；
禁词真源：skills/dealer-ops-suite/assets/copy/对外文案.json 的 l1_banned / expiry_terms。
钉死六类回归中的「对外文案防回潮」：A1 禁词、A2 内部泄漏、A3 效期开关、
A4 文档属性、A5 冻结表基线比对、A6 附表编号与交叉引用，以及 CLI 退出行为。
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from check_public_copy import clean_properties, load_copy, scan  # noqa: E402

# references/对外文案治理.md「关闭时的硬断言」清单：全簿命中数必须为 0
DOC89_TERMS = ("效期", "批次", "到期", "临期", "预警", "剩余月数", "FIFO")


def save_book(path, sheets, *, creator=None, last_modified_by=None):
    """openpyxl 默认会写入 creator=openpyxl；默认清空作者痕迹，便于按需单独测 A4。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.properties.creator = "" if creator is None else creator
    wb.properties.lastModifiedBy = "" if last_modified_by is None else last_modified_by
    wb.save(path)
    return wb


class PublicCopyGateTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory(prefix=".test-public-copy-", dir=ROOT)
        self.base = Path(self._tmp.name)
        self.copy = load_copy()

    def tearDown(self):
        self._tmp.cleanup()

    def minimal_cfg(self, *, expiry=False, internal=(), frozen=(), sheet_map=None):
        """scan 只读 koujing.decisions 的四个键；按用例需要给最小口径。"""
        decisions = {"expiry": {"enabled": expiry},
                     "delta": {"frozen_sheets": list(frozen)}}
        if sheet_map is not None:
            decisions["sheet_map"] = sheet_map
        return {"koujing": {"internal_only": list(internal), "decisions": decisions}}

    def book(self, name, sheets, **props):
        path = self.base / name
        save_book(str(path), sheets, **props)
        return str(path)

    def clean_sheets(self):
        """含「甲方期初数量」表头与 08 附表的最小干净工作簿（全部为 l3 允许词）。"""
        return {
            "00-对账说明": [["对账说明"],
                           ["甲方期初数量", "期间发货", "乙方使用", "甲方期末"]],
            "08-试用装与中小样": [["类型", "甲方期初", "期间发货", "乙方使用", "甲方期末"],
                                 ["试用装", 0, 0, 0, 0]],
        }

    def test_clean_workbook_passes_scan(self):
        path = self.book("干净.xlsx", self.clean_sheets())
        result = scan(path, cfg=self.minimal_cfg(internal=["只许内部看的备注"]))
        self.assertTrue(result["ok"], result["problems"])
        self.assertEqual(result["problems"], [])
        self.assertFalse(result["expiry_enabled"])
        self.assertEqual(result["frozen_sheets"], [])
        self.assertGreater(result["scanned_strings"], 0)

    def test_l1_banned_words_in_cell_are_flagged(self):
        cases = {
            "BIP": "甲方进销存系统代号",   # 防回潮的核心钉子：BIP 必须被机器拦截
            "ERP": "内部系统名",
            "主档": "内部文件名",
            "挂账表": "内部表名",
            "FIFO": "内部作业口径",
            "先进先出": "内部作业口径",
            "重跑": "内部操作说明",
            "甲方出库": "非滚动台账术语",
        }
        for word, label in cases.items():
            with self.subTest(word=word):
                path = self.book("%s.xlsx" % word, {"对账": [["说明"], [word]]})
                result = scan(path, cfg=self.minimal_cfg(), strict=False)
                self.assertFalse(result["ok"], result["problems"])
                a1 = [p for p in result["problems"] if p.startswith("A1")]
                self.assertTrue(any(word in p for p in a1), result["problems"])
                self.assertTrue(any(label in p for p in a1), result["problems"])

    def test_banned_word_hidden_in_footer_comment_and_cell_is_counted(self):
        wb = save_book(str(self.base / "藏词.xlsx"),
                       {"对账": [["正常文案"], ["又一句正常文案", "BIP"]]})
        ws = wb["对账"]
        ws["B2"].comment = Comment("BIP 备注痕迹", "测试")
        ws.oddFooter.right.text = "BIP 尾注痕迹"
        path = self.base / "藏词.xlsx"
        wb.save(path)
        result = scan(str(path), cfg=self.minimal_cfg(), strict=False)
        self.assertFalse(result["ok"])
        a1 = [p for p in result["problems"] if p.startswith("A1")]
        self.assertEqual(len(a1), 1, result["problems"])
        self.assertIn("命中 3 处", a1[0])

    def test_internal_only_leak_is_flagged_and_clean_book_passes(self):
        secret = "内部结算口径备注"
        leaked = self.book("泄漏.xlsx", {"对账": [["说明"], ["口径：" + secret]]})
        result = scan(leaked, cfg=self.minimal_cfg(internal=[secret]), strict=False)
        self.assertFalse(result["ok"])
        self.assertTrue(any(p.startswith("A2") for p in result["problems"]),
                        result["problems"])
        clean = self.book("干净A2.xlsx", {"对账": [["说明"], ["口径：按合同约定执行"]]})
        result = scan(clean, cfg=self.minimal_cfg(internal=[secret]), strict=False)
        self.assertTrue(result["ok"], result["problems"])

    def test_expiry_terms_blocked_only_when_expiry_disabled(self):
        path = self.book("效期.xlsx", {"对账": [["说明"], ["剩余效期与到期提醒"]]})
        off = scan(path, cfg=self.minimal_cfg(expiry=False), strict=False)
        self.assertFalse(off["ok"])
        a3 = [p for p in off["problems"] if p.startswith("A3")]
        self.assertTrue(a3, off["problems"])
        self.assertIn("效期", a3[0])
        on = scan(path, cfg=self.minimal_cfg(expiry=True), strict=False)
        self.assertTrue(on["ok"], on["problems"])
        self.assertTrue(on["expiry_enabled"])

    def test_doc_properties_author_trace_flagged_then_cleaned(self):
        path = self.book("署名.xlsx", {"对账": [["说明"], ["正常文案"]]},
                         creator="张三", last_modified_by="李四")
        result = scan(str(path), cfg=self.minimal_cfg(), strict=False)
        self.assertFalse(result["ok"])
        a4 = [p for p in result["problems"] if p.startswith("A4")]
        self.assertTrue(any("dc:creator" in p for p in a4), result["problems"])
        self.assertTrue(any("cp:lastModifiedBy" in p for p in a4), result["problems"])
        clean_properties(str(path))
        result = scan(str(path), cfg=self.minimal_cfg(), strict=False)
        self.assertTrue(result["ok"], result["problems"])
        self.assertFalse([p for p in result["problems"] if p.startswith("A4")])

    def test_frozen_sheet_cell_diff_against_baseline_is_flagged(self):
        sheets = {"01-对账单": [["甲方", "乙方"], [1, 2]],
                  "02-库存差异表": [["商品", "差异"], ["甲", 0]]}
        baseline = self.book("基线.xlsx", sheets)
        cfg = self.minimal_cfg(frozen=["01"])
        identical = self.book("未变.xlsx", sheets)
        result = scan(identical, cfg=cfg, baseline=baseline, strict=False)
        self.assertTrue(result["ok"], result["problems"])
        self.assertEqual(result["frozen_sheets"], ["01"])
        moved = dict(sheets)
        moved["01-对账单"] = [["甲方", "乙方"], [1, 999]]
        changed_frozen = self.book("冻结被改.xlsx", moved)
        result = scan(changed_frozen, cfg=cfg, baseline=baseline, strict=False)
        self.assertFalse(result["ok"])
        a5 = [p for p in result["problems"] if p.startswith("A5")]
        self.assertTrue(any("01-对账单" in p and "B2" in p for p in a5), result["problems"])
        touched = dict(sheets)
        touched["02-库存差异表"] = [["商品", "差异"], ["乙", 5]]
        changed_free = self.book("非冻结被改.xlsx", touched)
        result = scan(changed_free, cfg=cfg, baseline=baseline, strict=False)
        self.assertTrue(result["ok"], result["problems"])

    def test_sheet_map_targets_and_dangling_references_are_checked(self):
        sheet_map = {"08": "试用装与中小样", "09": "效期预警"}
        cfg = self.minimal_cfg(sheet_map=sheet_map)
        missing = self.book("缺08.xlsx", {"00-对账说明": [["对账说明"]]})
        result = scan(missing, cfg=cfg, strict=False)
        self.assertTrue(any(p.startswith("A6") and "08-试用装与中小样" in p
                            for p in result["problems"]), result["problems"])
        complete = {"00-对账说明": [["对账说明"]],
                    "08-试用装与中小样": [["类型"], ["试用装"]]}
        ok_book = self.book("齐.xlsx", complete)
        result = scan(ok_book, cfg=cfg, strict=False)
        self.assertFalse([p for p in result["problems"] if p.startswith("A6")],
                         result["problems"])
        dangling = dict(complete)
        dangling["00-对账说明"] = [["明细见附表 09"]]
        result = scan(self.book("悬空.xlsx", dangling), cfg=cfg, strict=False)
        self.assertTrue(any(p.startswith("A6") and "附表 09" in p
                            for p in result["problems"]), result["problems"])
        legacy = dict(complete)
        legacy["09-效期预警"] = [["预警档"], ["正常"]]
        result = scan(self.book("多09.xlsx", legacy), cfg=cfg, strict=False)
        self.assertTrue(any(p.startswith("A6") and "效期工作表" in p
                            for p in result["problems"]), result["problems"])

    def test_strict_mode_exits_1_and_clean_book_returns_result(self):
        path = self.book("违禁.xlsx", {"对账": [["说明"], ["BIP"]]})
        with self.assertRaises(SystemExit) as ctx:
            scan(path, cfg=self.minimal_cfg(), strict=True)
        self.assertEqual(ctx.exception.code, 1)
        clean = self.book("通过.xlsx", self.clean_sheets())
        result = scan(clean, cfg=self.minimal_cfg(), strict=True)
        self.assertTrue(result["ok"], result["problems"])

    def test_copy_asset_pins_bip_and_hard_assertion_terms(self):
        # 对外文案.json 是唯一文案真源：BIP 防回潮条目与 md:89 硬断言清单不得被删
        self.assertTrue(any(item["pattern"] == "BIP" and item["label"] == "甲方进销存系统代号"
                            for item in self.copy["l1_banned"]))
        for term in DOC89_TERMS:
            if term == "FIFO":
                self.assertIn(term, [item["pattern"] for item in self.copy["l1_banned"]])
            else:
                self.assertIn(term, self.copy["expiry_terms"], term)

    def run_cli(self, *args):
        return subprocess.run([sys.executable, str(SCRIPTS / "check_public_copy.py"), *args],
                              cwd=ROOT, text=True, capture_output=True, timeout=30)

    def write_cfg_file(self, cfg, name="cfg.json"):
        path = self.base / name
        path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        return str(path)

    def test_cli_gate_fails_closed_on_banned_word(self):
        path = self.book("违禁CLI.xlsx", {"对账": [["说明"], ["BIP"]]})
        cfg = self.write_cfg_file(self.minimal_cfg())
        run = self.run_cli("--book", path, "--cfg", cfg)
        self.assertEqual(run.returncode, 1, run.stdout + "\n" + run.stderr)
        self.assertIn("[对外门禁未通过]", run.stderr)
        self.assertIn("BIP", run.stderr)
        self.assertEqual(run.stdout, "")

    def test_cli_json_mode_reports_problems_without_exit(self):
        path = self.book("违禁JSON.xlsx", {"对账": [["说明"], ["BIP"]]})
        cfg = self.write_cfg_file(self.minimal_cfg())
        run = self.run_cli("--book", path, "--cfg", cfg, "--json")
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        payload = json.loads(run.stdout)
        self.assertFalse(payload["ok"])
        self.assertTrue(any(p.startswith("A1") for p in payload["problems"]))
        self.assertEqual(payload["book"], os.path.abspath(path))

    def test_cli_clean_book_passes_and_clean_props_clears_author(self):
        path = self.book("干净CLI.xlsx", self.clean_sheets())
        cfg = self.write_cfg_file(self.minimal_cfg(
            internal=["只许内部看的备注"],
            sheet_map={"08": "试用装与中小样", "09": "效期预警"}))
        run = self.run_cli("--book", path, "--cfg", cfg)
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        self.assertTrue(json.loads(run.stdout)["ok"])
        signed = self.book("署名CLI.xlsx", {"对账": [["说明"], ["正常文案"]]},
                           creator="张三", last_modified_by="李四")
        run = self.run_cli("--book", signed, "--clean-props")
        self.assertEqual(run.returncode, 0, run.stdout + "\n" + run.stderr)
        self.assertIn("OK", run.stdout)
        result = scan(signed, cfg=self.minimal_cfg(), strict=False)
        self.assertTrue(result["ok"], result["problems"])


if __name__ == "__main__":
    unittest.main()
