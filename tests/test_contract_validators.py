#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合同线纯函数级校验与 spec 完整性回归（unittest 风格）。

收纳自原仓 test_validators.py（信用代码 / 合同编号 / 日期 / 金额大写 /
输入过滤 / 计划指纹）与 test_specs.py（spec 唯一机器事实源的一致性），
并按合并仓裁决改写一处：

  公开版 fc.json 的 locked_content 为空表（T008 脱敏裁决）——空表时
  verify_fill 的锁定条款表校验是 no-op，必须显式披露「未登记」而不是
  静默报「完整」。因此原仓「fc 锁定表登记 2 条」的断言改写为：
  locked_content == [] + verify_fill.py 源码含「未登记」披露文案。

  品牌脱敏（原仓品牌码与品牌名已全部替换为 DM / 通用占位）：合并仓
  build_plan 硬校验品牌码 DM（check_contract_no_consistency 对非 DM
  品牌码给 warning），全部编号样例已 DM 化；「未知品牌仅警告」用例
  保留 AB 前缀，恰是 DM 规则的测试。
"""
import json
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "dealer-ops-suite" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from build_plan import (check_amount_cn, check_contract_no,  # noqa: E402
                        check_contract_no_consistency, check_date,
                        check_phone, check_uscc, num2cn, validate_one)
from plan_trust import canon, check_plan, check_plan_selfconsistency, values_digest  # noqa: E402

SPEC_DIR = SCRIPTS / "spec"
NAMES = ["jx", "dx", "fc"]

KNOWN_VTYPES = {
    "text", "contract_no", "company", "uscc", "person", "phone", "address",
    "region", "channel", "date", "percent_rebate", "percent_any",
    "amount_num", "amount_cn", "amount_wan", "day", "int",
}


def load_spec(name):
    with open(SPEC_DIR / f"{name}.json", encoding="utf-8") as fh:
        return json.load(fh)


def all_items(spec):
    return (list(spec.get("fields", [])) + list(spec.get("cells", []))
            + list(spec.get("grids", [])))


class ValidatorTest(unittest.TestCase):
    """纯函数级校验：信用代码、日期、金额大写、输入过滤、计划指纹。"""

    # ---------------------------------------------------------- 信用代码
    def test_uscc_accepts_valid_code(self):
        ok, msg = check_uscc("91310101MA1FL2XX71")
        self.assertTrue(ok, msg)

    def test_uscc_rejects_wrong_check_digit(self):
        ok, _ = check_uscc("91310101MA1FL2XX72")
        self.assertFalse(ok)

    def test_uscc_rejects_forbidden_letters(self):
        # I / O / Z / S / V 不参与信用代码编码
        for ch in "IOZSV":
            with self.subTest(ch=ch):
                ok, _ = check_uscc("91310101MA1FL2XX7" + ch)
                self.assertFalse(ok)

    def test_uscc_rejects_wrong_length(self):
        self.assertFalse(check_uscc("91310101MA1FL2XX7")[0])

    # ---------------------------------------------------------- 合同编号格式
    def test_contract_no_accepts(self):
        for code in (
            "DM-JX-2026-001",
            "DM-DX-2026-999",
            "DM-JX-JS-2026-001",   # 带扩展段
            "待定",                 # 起草阶段占位
        ):
            with self.subTest(code=code):
                ok, msg = check_contract_no(code)
                self.assertTrue(ok, f"{code}: {msg}")

    def test_contract_no_rejects(self):
        for code in (
            "JX-2026-001",          # 缺品牌码
            "dm-jx-2026-001",       # 小写
            "DM_JX_2026_001",       # 下划线
            "DM-JX-2026-01",        # 序号不足三位
            "DM-JX-2026-000",       # 序号自 001 起
            "DM－JX－2026－001",    # 全角连字符
            "示例-JX-2026-001",     # 中文品牌码（脱敏占位；中文不得进入编号）
            "DM-JX-2026-0001",      # 序号超三位
        ):
            with self.subTest(code=code):
                ok, _ = check_contract_no(code)
                self.assertFalse(ok, code)

    # ---------------------------------------------------------- 编号跨字段一致性
    def test_contract_no_type_code_must_match_contract(self):
        errors, _ = check_contract_no_consistency("DM-DX-2026-001", "JX", "2026-10-20")
        self.assertTrue(errors)
        self.assertIn("类型码", errors[0])

    def test_contract_no_year_must_match_sign_date(self):
        errors, _ = check_contract_no_consistency("DM-JX-2025-001", "JX", "2026-10-20")
        self.assertTrue(errors)
        self.assertIn("年份", errors[0])

    def test_contract_no_unknown_brand_warns_only(self):
        # 品牌 AB 未登记：只警告不阻断（合并仓登记的品牌码为 DM）
        errors, warns = check_contract_no_consistency("AB-JX-2026-001", "JX", "2026-10-20")
        self.assertFalse(errors)
        self.assertTrue(warns)

    def test_contract_no_consistent_case_is_clean(self):
        self.assertEqual(
            check_contract_no_consistency("DM-JX-2026-001", "JX", "2026-10-20"),
            ([], []))

    def test_contract_no_placeholder_skips_cross_checks(self):
        self.assertEqual(
            check_contract_no_consistency("待定", "JX", "2026-10-20"),
            ([], []))

    # ---------------------------------------------------------- 日期与电话
    def test_date_accepts_leap_day(self):
        self.assertTrue(check_date("2024-02-29")[0])

    def test_date_rejects_impossible_dates(self):
        for bad in ("2026-02-30", "2026-13-01", "2026-00-10", "2026-1-1"):
            with self.subTest(bad=bad):
                self.assertFalse(check_date(bad)[0])

    def test_phone_rules(self):
        self.assertTrue(check_phone("13800138000")[0])
        self.assertTrue(check_phone("020-88886666")[0])
        self.assertFalse(check_phone("1380013800")[0])

    # ---------------------------------------------------------- 金额大写
    def test_num2cn(self):
        for num, expected in (
            (0, "零"),
            (5, "伍"),
            (10, "壹拾"),
            (10001, "壹万零壹"),
            (50010, "伍万零壹拾"),
            (100000, "壹拾万"),
        ):
            with self.subTest(num=num):
                self.assertEqual(num2cn(num), expected)

    def test_amount_cn_must_match_number(self):
        self.assertTrue(check_amount_cn("壹万元", 10000)[0])
        self.assertFalse(check_amount_cn("贰万元", 10000)[0])

    # ---------------------------------------------------------- 输入过滤
    def test_validate_one_rejects_invisible_characters(self):
        field = {"vtype": "text", "required": True}
        for bad in ("abc\u200b", "abc\u202e", "\ufeffabc", "a\tb"):
            with self.subTest(bad=repr(bad)):
                ok, _ = validate_one(field, bad)
                self.assertFalse(ok)

    def test_validate_one_rejects_unknown_vtype(self):
        ok, msg = validate_one({"vtype": "no_such_type", "required": True}, "x")
        self.assertFalse(ok)
        self.assertIn("未知的校验类型", msg)

    def test_validate_one_requires_value_for_required_field(self):
        self.assertFalse(validate_one({"vtype": "text", "required": True}, "")[0])
        self.assertTrue(validate_one({"vtype": "text", "required": False}, "")[0])

    # ---------------------------------------------------------- 计划指纹
    def test_values_digest_is_order_independent(self):
        self.assertEqual(values_digest({"a": 1, "b": 2}), values_digest({"b": 2, "a": 1}))
        self.assertNotEqual(values_digest({"a": 1}), values_digest({"a": 2}))

    def test_canon_is_key_order_independent(self):
        self.assertEqual(canon({"x": 1, "y": [1, 2]}), canon({"y": [1, 2], "x": 1}))

    def test_check_plan_fails_closed_without_roots(self):
        """缺 spec / values / 模板时，可信性检查必须直接失败，不能返回空列表放行。"""
        self.assertTrue(check_plan({}, None, None, None))
        self.assertTrue(check_plan({"ops": []}, "/tmp/x.json", None, None))

    def test_plan_selfconsistency_detects_missing_expect(self):
        plan = {"ops": [{"field_id": "A", "value": "1"}]}
        self.assertTrue(check_plan_selfconsistency(plan))

    def test_plan_selfconsistency_is_type_strict(self):
        """1 与 "1" 不能算一致，否则类型混淆可以绕过。"""
        plan = {"ops": [{"field_id": "A", "value": "1", "expect": 1}]}
        self.assertTrue(check_plan_selfconsistency(plan))


class SpecIntegrityTest(unittest.TestCase):
    """字段规格与配套文档的一致性。

    spec 是唯一机器事实源，人读文档必须与它对齐；示例值本身也必须能通过校验，
    否则用户照抄示例就会被 build_plan 拒掉（第三轮审计实证过这个问题）。
    """

    def test_ids_are_unique_and_keys_present(self):
        for name in NAMES:
            with self.subTest(name=name):
                spec = load_spec(name)
                ids = [x["id"] for x in all_items(spec)]
                self.assertTrue(ids, f"{name}.json 没有任何字段")
                dup = {i for i in ids if ids.count(i) > 1}
                self.assertFalse(dup, f"{name}.json 字段 ID 重复：{dup}")
                for item in spec.get("fields", []) + spec.get("cells", []):
                    self.assertTrue(item["id"].startswith(spec["doc_key"] + "-"), item["id"])
                    self.assertIsInstance(item.get("required"), bool, item["id"])
                    self.assertTrue(item.get("label"), item["id"])
                    self.assertIn(item.get("vtype"), KNOWN_VTYPES,
                                  f"{item['id']} vtype={item.get('vtype')}")

    def test_field_targets_are_well_formed(self):
        for name in NAMES:
            with self.subTest(name=name):
                spec = load_spec(name)
                for f in spec.get("fields", []):
                    target = f.get("target") or {}
                    self.assertTrue(target.get("anchor"), f"{f['id']} 缺少锚点")
                    self.assertIn(f.get("mode"), ("append", "regex"), f["id"])
                    if f["mode"] == "regex":
                        self.assertTrue(f.get("pattern"), f"{f['id']} regex 模式缺少 pattern")
                        # 正则必须能编译，避免运行期才炸
                        re.compile(f["pattern"])
                    self.assertIn("{", f.get("render", ""), f"{f['id']} render 不含占位符")
                for c in spec.get("cells", []):
                    self.assertIsInstance(c.get("table"), int, c["id"])
                    self.assertIsInstance(c.get("row"), int, c["id"])
                    self.assertIsInstance(c.get("col"), int, c["id"])
                    self.assertIn("{", c.get("render", ""), f"{c['id']} render 不含占位符")

    def test_examples_pass_own_validation(self):
        """示例值必须自己就能通过校验 —— 用户是照着它填的。

        全量不变式，无豁免：region 字段示例已修为「示例」
        （脱敏曾写成「示例省/示例市」，被 check_region 按设计拒绝——
        迁移测试时暴露，spec 已同步修正）。任何示例不过校验都判失败。
        """
        for name in NAMES:
            with self.subTest(name=name):
                spec = load_spec(name)
                bad = []
                for item in spec.get("fields", []) + spec.get("cells", []):
                    example = item.get("example")
                    if not example:
                        continue
                    ok, msg = validate_one(item, example)
                    if not ok:
                        bad.append(f"{item['id']} {item['label']}：{example!r} → {msg}")
                self.assertFalse(
                    bad, "示例值不合法：\n" + "\n".join(bad))

    def test_field_map_covers_every_id(self):
        """field-map.md 是人读版本，必须覆盖 spec 的每一个 ID。"""
        with open(SCRIPTS.parent / "references" / "field-map.md", encoding="utf-8") as fh:
            doc = fh.read()
        for name in NAMES:
            with self.subTest(name=name):
                missing = [x["id"] for x in all_items(load_spec(name)) if x["id"] not in doc]
                self.assertFalse(missing, f"field-map.md 缺少：{missing}")

    def test_readonly_declared_and_locked_public_empty(self):
        """jx/dx 须登记 ≥4 条甲方固定信息；fc 的锁定表在公开版为空表（T008）。

        空表意味着 verify_fill 的锁定条款表校验本次是 no-op，源码里必须
        显式披露「未登记」，不能静默报「完整」。
        """
        for name in ("jx", "dx"):
            with self.subTest(name=name):
                self.assertGreaterEqual(
                    len(load_spec(name).get("readonly_check") or []), 4, name)
        fc = load_spec("fc")
        self.assertEqual(fc.get("locked_content"), [])
        verify_fill_src = (SCRIPTS / "verify_fill.py").read_text(encoding="utf-8")
        self.assertIn("未登记", verify_fill_src)

    def test_fc_dates_are_required_and_documented(self):
        """FC 两个签署日期是必填；三份人读文档必须同步，否则流程会卡死。"""
        fc = load_spec("fc")
        dates = {c["id"]: c for c in fc["cells"] if c["id"] in ("FC-C1", "FC-C2")}
        self.assertEqual(len(dates), 2)
        for cid, cell in dates.items():
            self.assertIs(cell["required"], True, cid)

        docs = {}
        for name in ("field-map.md", "confirm-template.md", "validation-rules.md"):
            with open(SCRIPTS.parent / "references" / name, encoding="utf-8") as fh:
                docs[name] = fh.read()
        self.assertIn("附件签署日期", docs["confirm-template.md"], "确认清单漏问附件日期")
        self.assertIn("必填", docs["field-map.md"])
        self.assertIn("不自动推断", docs["validation-rules.md"])


if __name__ == "__main__":
    unittest.main()
