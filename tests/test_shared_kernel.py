#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""共享内核钉死测试：locate_column / normalize_barcode / sha256_file golden。

断言全部来自探针脚本（/tmp）对现有实现的实测输出，钉死的是收敛前行为基线：
- locate_column 对齐 monthly_recon_engine.column() 的最严语义；
- normalize_barcode 对齐 settlement_rules.barcode_of() 的权威语义；
- sha256_file 的算法摘要用硬编码 golden 值钉死（不得在测试内现算）。
"""

import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "skills", "dealer-ops-suite", "scripts"))

from schema_common import (  # noqa: E402
    ColumnAmbiguous,
    ColumnLookupError,
    ColumnMissing,
    locate_column,
    normalize_barcode,
    sha256_file,
)
from schema_common import FieldColumnConflict, map_field_columns  # noqa: E402

GOLDEN_INPUT = b"dealer-ops-suite-kernel"
GOLDEN_SHA256 = "e087ec7afd55f9c3987dc08222fd259b74e660c1a20e69dbcf3e7d5ce7ec57c0"


class Sha256GoldenTest(unittest.TestCase):
    def test_fixed_bytes_match_hardcoded_digest(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "golden.bin")
            with open(path, "wb") as fh:
                fh.write(GOLDEN_INPUT)
            self.assertEqual(sha256_file(path), GOLDEN_SHA256)


class LocateColumnTest(unittest.TestCase):
    def test_exact_hit_returns_zero_based_index(self):
        self.assertEqual(locate_column(["条码", "商品名称"], ("商品条码", "条码")), 0)

    def test_first_alias_wins_on_exact_round(self):
        self.assertEqual(locate_column(["条码", "商品条码"], ("商品条码", "条码")), 1)

    def test_substring_fallback_hits_containing_header(self):
        self.assertEqual(locate_column(["退货商品条码"], ("商品条码",)), 0)

    def test_generalized_token_only_accepts_full_header(self):
        self.assertEqual(locate_column(["数量"], ("数量",)), 0)
        # 泛化词跳过子串轮：「发货数量」不被「数量」命中，缺失即阻断
        with self.assertRaises(ColumnMissing):
            locate_column(["发货数量"], ("数量",))
        self.assertEqual(locate_column(["签收数量"], ("签收数量",)), 0)

    def test_exact_multiple_hits_raise_ambiguous(self):
        with self.assertRaises(ColumnAmbiguous) as ctx:
            locate_column(["数量", "数量"], ("数量",))
        self.assertEqual(ctx.exception.hits, [0, 1])
        self.assertEqual(ctx.exception.alias, "数量")
        self.assertEqual(ctx.exception.headers, ["数量", "数量"])
        self.assertIsInstance(ctx.exception, ColumnLookupError)

    def test_substring_multiple_hits_raise_ambiguous(self):
        with self.assertRaises(ColumnAmbiguous) as ctx:
            locate_column(["商品条码A", "商品条码B"], ("商品条码",))
        self.assertEqual(ctx.exception.hits, [0, 1])
        self.assertIn("不唯一", str(ctx.exception))

    def test_missing_required_raises_with_aliases_and_sheet_name(self):
        with self.assertRaises(ColumnMissing) as ctx:
            locate_column(["条码"], ("单价", "供货价"), sheet_name="销售")
        self.assertEqual(ctx.exception.aliases, ["单价", "供货价"])
        self.assertEqual(ctx.exception.sheet_name, "销售")
        self.assertEqual(ctx.exception.headers, ["条码"])
        message = str(ctx.exception)
        self.assertIn("单价／供货价", message)
        self.assertIn("工作表「销售」", message)
        self.assertIn("条码", message)

    def test_missing_optional_returns_none(self):
        self.assertIsNone(locate_column(["条码"], ("单价", "供货价"), required=False))

    def test_cn_parenthesized_content_is_ignored(self):
        self.assertEqual(locate_column(["商品名称（含赠品）", "单价"], ("商品名称",)), 0)
        # 整个表头都在括号内 → 归一化后为空，不参与匹配
        with self.assertRaises(ColumnMissing):
            locate_column(["（备注）", "单价"], ("备注",))

    def test_whitespace_and_paren_alias_normalize_like_monthly(self):
        self.assertEqual(locate_column(["签 收 数 量"], ("签收数量",)), 0)
        self.assertEqual(locate_column(["单价"], ("单价(元)",)), 0)


class NormalizeBarcodeTest(unittest.TestCase):
    def test_equivalence_classes_match_settlement_barcode_of(self):
        cases = {
            "6901234567890": "6901234567890",     # 13 位数字串
            1300000000001.0: "1300000000001",     # 整值 float
            "1300000000001.0": "1300000000001",   # ".0" 尾巴剥掉
            1300000000001: "1300000000001",       # int
            None: "",
            "  6901234567890  ": "6901234567890",  # 首尾空白
            "ABC-123": "ABC-123",                 # 非数字字符串透传
            12.5: "12.5",                         # 非整值 float
            0.0: "0",
            "690123456789.0": "690123456789",
            True: "True",                         # bool 走 str() 分支（与 barcode_of 一致）
            "": "",
        }
        for value, expected in cases.items():
            self.assertEqual(normalize_barcode(value), expected, "input=%r" % (value,))

    def test_result_is_stripped_str(self):
        self.assertIsInstance(normalize_barcode(1300000000001), str)
        self.assertEqual(normalize_barcode(" 0.0 "), "0")


class MapFieldColumnsTest(unittest.TestCase):
    """map_field_columns 钉死用例：语义与 check_policy.map_columns 收敛前的别名循环逐字一致。"""

    ALIASES = {
        "barcode": ("商品条码", "条码"),
        "name": ("商品名称", "品名"),
        "price": ("结算价", "合同价"),
        "start": ("生效起", "生效日期"),
        "end": ("生效止", "失效日期"),
    }

    def test_exact_and_substring_hits_are_zero_based(self):
        mapping, notes = map_field_columns(
            ["商品条码", "商品名称", "结算价", "生效起", "生效止"], dict(self.ALIASES))
        self.assertEqual(mapping, {"barcode": 0, "name": 1, "price": 2, "start": 3, "end": 4})
        self.assertEqual(notes, [])

    def test_substring_match_respects_max_suffix_boundary(self):
        # len(表头) = len(别名)+5 命中；+6 不命中（现行 +5 上限逐字保留）
        self.assertEqual(map_field_columns(["结算价一览表"], {"price": ("价",)}),
                         ({"price": 0}, []))
        self.assertEqual(map_field_columns(["一次性结算价格补差说明"], {"price": ("价",)}),
                         ({}, []))
        # max_suffix=0 退化为仅全等
        self.assertEqual(map_field_columns(["商品条码栏"], {"barcode": ("条码",)}, max_suffix=0),
                         ({}, []))
        self.assertEqual(map_field_columns(["条码"], {"barcode": ("条码",)}, max_suffix=0),
                         ({"barcode": 0}, []))

    def test_empty_and_none_headers_are_skipped(self):
        mapping, notes = map_field_columns([None, "", "条码"], {"barcode": ("条码",)})
        self.assertEqual(mapping, {"barcode": 2})
        self.assertEqual(notes, [])

    def test_note_keyword_precedes_field_match(self):
        # 「备注」既是 note 关键词又是 barcode 别名：必须计入 notes，不参与字段匹配
        mapping, notes = map_field_columns(["条码", "备注"], {"barcode": ("条码", "备注")},
                                           note_keywords=("备注",))
        self.assertEqual(mapping, {"barcode": 0})
        self.assertEqual(notes, [1])

    def test_first_column_wins_for_duplicate_alias_no_conflict(self):
        # 同一 field 已命中即跳过：同义两列不阻断（与现行 `if field in mapping: continue` 一致）
        mapping, notes = map_field_columns(["条码", "商品条码"], {"barcode": ("商品条码", "条码")})
        self.assertEqual(mapping, {"barcode": 0})
        self.assertEqual(notes, [])

    def test_assigned_preoccupation_conflicts(self):
        with self.assertRaises(FieldColumnConflict) as ctx:
            map_field_columns(["条码", "结算价"], {"barcode": ("条码",), "price": ("结算价",)},
                              assigned={0: "price"})
        self.assertEqual(ctx.exception.header, "条码")
        self.assertEqual(ctx.exception.left_field, "price")
        self.assertEqual(ctx.exception.right_field, "barcode")
        self.assertEqual(str(ctx.exception), "表头「条码」同时映射为 price 与 barcode，请修正表头")

    def test_same_column_hitting_two_fields_conflicts_without_break(self):
        # 命中后不 break：同列再命中后续字段必须阻断（现行循环逐字语义）
        with self.assertRaises(FieldColumnConflict) as ctx:
            map_field_columns(["商品条码"], {"barcode": ("商品条码",), "name": ("商品",)})
        self.assertEqual(ctx.exception.header, "商品条码")
        self.assertEqual(ctx.exception.left_field, "barcode")
        self.assertEqual(ctx.exception.right_field, "name")
        self.assertEqual(str(ctx.exception), "表头「商品条码」同时映射为 barcode 与 name，请修正表头")

    def test_conflict_is_a_column_lookup_error(self):
        self.assertTrue(issubclass(FieldColumnConflict, ColumnLookupError))
        exc = FieldColumnConflict("msg", "表头", "left", "right", headers=["表头"])
        self.assertEqual(exc.headers, ["表头"])

    def test_assigned_on_unmatched_column_is_ignored(self):
        mapping, notes = map_field_columns(["条码", "无关列"], {"barcode": ("条码",)},
                                           assigned={1: "decal"})
        self.assertEqual(mapping, {"barcode": 0})
        self.assertEqual(notes, [])

    def test_pre_assigned_field_is_skipped_for_its_own_column(self):
        # 价格列已由 rate_value_column 定位并预占用：别名循环不再重复认领（不报 price vs price），
        # 其余字段照常映射——与现行 `if field in mapping: continue` 编排语义一致
        aliases = {"barcode": ("条码",), "price": ("结算价",)}
        mapping, notes = map_field_columns(["条码", "结算价"], aliases, assigned={1: "price"})
        self.assertEqual(mapping, {"barcode": 0})
        self.assertEqual(notes, [])
        # 同一列被其他字段命中仍阻断
        with self.assertRaises(FieldColumnConflict) as ctx:
            map_field_columns(["条码", "结算价"], {"barcode": ("条码",), "name": ("结算价",)},
                              assigned={1: "price"})
        self.assertEqual(ctx.exception.left_field, "price")
        self.assertEqual(ctx.exception.right_field, "name")


if __name__ == "__main__":
    unittest.main()
