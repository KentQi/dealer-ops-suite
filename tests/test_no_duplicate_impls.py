"""共享内核去重门禁：防止列定位／条码归一化／金额舍入／SHA-256 的本地实现回归。

对应 GitHub issue #2 #3 #4 #5 的验收项「机制性去重门禁」：
共享实现收敛到 schema_common / settlement_rules 后，任何脚本不得再私养一份。
只做源码文本断言，不执行任何引擎，也不依赖运行目录。
"""
import os
import unittest

_SCRIPTS = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "skills", "dealer-ops-suite", "scripts"))


def _read(name):
    with open(os.path.join(_SCRIPTS, name), encoding="utf-8") as fh:
        return fh.read()


class Sha256ConvergenceTest(unittest.TestCase):
    """issue #5：文件 SHA-256 只允许 schema_common.sha256_file 一份。"""

    def test_canonical_lives_in_schema_common(self):
        self.assertIn("def sha256_file", _read("schema_common.py"))

    def test_engines_have_no_local_sha(self):
        for name in ("recon_engine.py", "monthly_recon_engine.py"):
            body = _read(name)
            self.assertNotIn("hashlib.sha256", body, f"{name} 不得私养 SHA-256 计算")
            self.assertNotIn("def _source_sha256", body)
            self.assertNotIn("def workbook_sha256_impl", body)

    def test_review_tools_import_shared_impl(self):
        for name in ("review_common.py", "contract_release.py"):
            body = _read(name)
            self.assertNotIn("def sha256_file", body, f"{name} 的本地 sha256_file 必须删除")
            self.assertIn("sha256_file", body)

    def test_contract_release_keeps_findings_digest_only(self):
        # findings_digest 是 findings JSON 的结构化摘要，语义不同，允许保留 hashlib。
        body = _read("contract_release.py")
        self.assertIn("def findings_digest", body)


class BarcodeConvergenceTest(unittest.TestCase):
    """issue #3：条码归一化只允许 schema_common.normalize_barcode 一份。"""

    def test_canonical_lives_in_schema_common(self):
        self.assertIn("def normalize_barcode", _read("schema_common.py"))

    def test_settlement_rules_reexports_authority(self):
        body = _read("settlement_rules.py")
        self.assertNotIn("def barcode_of", body, "barcode_of 实现体必须删除，只留再导出")
        self.assertIn("normalize_barcode as barcode_of", body)

    def test_monthly_delegates(self):
        body = _read("monthly_recon_engine.py")
        self.assertIn("normalize_barcode", body)
        self.assertNotIn('endswith(".0")', body, "月结不得保留本地 .0 剥尾逻辑")

    def test_check_policy_delegates(self):
        body = _read("check_policy.py")
        self.assertIn("normalize_barcode", body)


class ColumnConvergenceTest(unittest.TestCase):
    """issue #2：列定位收敛——引擎走 locate_column，政策走 map_field_columns + rate_value_column。"""

    def test_canonical_primitives_exist(self):
        body = _read("schema_common.py")
        self.assertIn("def locate_column", body)
        self.assertIn("def map_field_columns", body)

    def test_both_engines_use_locate_column(self):
        for name in ("recon_engine.py", "monthly_recon_engine.py"):
            self.assertIn("locate_column", _read(name), f"{name} 必须经共享 locate_column 取列")

    def test_check_policy_uses_shared_mappers(self):
        body = _read("check_policy.py")
        self.assertIn("map_field_columns", body)
        self.assertIn("rate_value_column", body)
        self.assertNotIn(
            "for field, aliases in COL_ALIASES.items()",
            body, "政策自检不得保留本地多字段别名循环")

    def test_recon_engine_has_no_first_hit_alias_loop(self):
        self.assertNotIn("for name in names", _read("recon_engine.py"),
                         "铺货引擎不得保留「先到先得」别名循环")


class MoneyConvergenceTest(unittest.TestCase):
    """issue #4：金额舍入只允许 settlement_rules.money 一份（Decimal ROUND_HALF_UP）。"""

    def test_canonical_lives_in_settlement_rules(self):
        body = _read("settlement_rules.py")
        self.assertIn("def money", body)
        self.assertIn("ROUND_HALF_UP", body)

    def test_writeback_quantizes_via_money(self):
        body = _read("writeback_ledger.py")
        self.assertIn("from settlement_rules import money", body)
        self.assertNotIn("round(number, 2)", body, "回写不得保留本地 half-even 舍入")

    def test_review_total_anchored_on_money(self):
        body = _read("review_settlement.py")
        self.assertIn("from settlement_rules import money", body)
        self.assertIn("money(", body)

    def test_recon_engine_uses_authority(self):
        body = _read("recon_engine.py")
        self.assertIn("money(", body)


class KernelPinningTestExistsTest(unittest.TestCase):
    """钉死测试（golden vector）必须随内核存在。"""

    def test_pinning_suite_present(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_shared_kernel.py")
        self.assertTrue(os.path.isfile(path), "tests/test_shared_kernel.py 钉死套件不得删除")


if __name__ == "__main__":
    unittest.main()
