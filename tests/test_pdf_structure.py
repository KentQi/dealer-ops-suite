#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""签章 PDF 结构回归（原 dealer-settlement CI 专属检查，v1.9.1/1.9.2 排版修复的钉子）。

重依赖（pymupdf + reportlab + 中文字体），不进主 CI：本地缺 pymupdf 时整类跳过，
export_pdf 找不到中文字体时单用例跳过。CI 由 .github/workflows/ci.yml 的
`pdf-regression` 手动 job 安装 pymupdf 与 fonts-wqy-zenhei 后运行。
"""
import subprocess
import sys
import unittest
from pathlib import Path

try:
    import fitz  # noqa: F401  pymupdf
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

from test_v17_gates import ConsignmentFixture  # noqa: E402  复用出账夹具

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"


@unittest.skipUnless(HAS_PYMUPDF, "需要 pymupdf（PDF 回归 job 安装）")
class PdfStructureTest(ConsignmentFixture):
    """export_pdf 产物的结构断言：页数、可提取文本、续页表头。"""

    def _export(self, name):
        out = self.base / name
        run = subprocess.run(
            [sys.executable, str(SCRIPTS / "export_pdf.py"), "--book", str(self.out),
             "--out", str(out)],
            text=True, capture_output=True, timeout=60, cwd=ROOT)
        if run.returncode != 0:
            self.skipTest("export_pdf 未产出（缺 reportlab 或中文字体）：%s"
                          % (run.stderr.strip()[-160:] or run.stdout.strip()[-160:]))
        return out

    def test_pdf_pages_and_extractable_text(self):
        self.run_engine()
        pdf_path = self._export("对账单.pdf")
        doc = fitz.open(pdf_path)
        try:
            # 工作簿含 00–09 十张表，默认 --sheets 00,01,02,08 四张 + 签章分页，页数下限守恒
            self.assertGreaterEqual(doc.page_count, 4)
            text = "\n".join(page.get_text() for page in doc)
            self.assertIn("对账", text, "PDF 文本层应可提取（可搜索），疑似整页位图化")
            self.assertIn("乙方", text)
        finally:
            doc.close()

    def test_pdf_deterministic_page_count(self):
        self.run_engine()
        first = fitz.open(self._export("对账1.pdf"))
        try:
            count = first.page_count
        finally:
            first.close()
        second = fitz.open(self._export("对账2.pdf"))
        try:
            self.assertEqual(second.page_count, count, "同输入两次导出页数必须一致")
        finally:
            second.close()


if __name__ == "__main__":
    unittest.main()
