"""共享内核去重门禁（AST 能力扫描）。

对应 GitHub issue #2 #3 #4 #5 的验收项「机制性去重门禁」。

与早期「固定文件清单 + 文本子串」版本不同（复核发现它拦不住新脚本私养
实现、也拦不住已覆盖文件里的措辞规避），本门禁：

1. 扫描 scripts/ 全目录——新脚本自动纳管，不依赖维护文件清单；
2. 按「能力」判定而非字面匹配：``import hashlib as h`` 别名、
   ``h.new("sha256")``、``hashlib.file_digest`` 等规避写法同样命中；
3. 保留结构断言：权威实现存在、调用方已接线、钉死套件在场。

能力 → 唯一权威实现：

===========  ===========================================================
文件 SHA-256  schema_common.sha256_file
条码归一化    schema_common.normalize_barcode（".0" 剥尾惯用法不得外流）
列定位        schema_common.locate_column / map_field_columns、
             settlement_rules.rate_value_column
金额舍入      settlement_rules.money（Decimal.quantize 不得越出
             settlement_rules）
===========  ===========================================================

违例签名（文件哈希能力）＝ SHA 计算 + 流式喂食惯用法（iter-lambda /
循环内 digest.update）或 hashlib.file_digest——即「重新实现流式文件哈希
算法」。对已经在内存中的字节串做一次 ``hashlib.sha256(raw)`` 属原语调用
（附件摘要、values_digest、findings_digest、待办 id 哈希、docx 逐部件
指纹），不是重复实现，不命中。exec/getattr 级动态规避不在 AST 防御范围，
靠 code review 兜底。
"""
import ast
import os
import tempfile
import unittest

_SCRIPTS = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "skills", "dealer-ops-suite", "scripts"))

# 每个能力的白名单 = 该能力唯一权威实现所在文件。
_SHA_WHITELIST = {"schema_common.py"}
_BARCODE_WHITELIST = {"schema_common.py"}
_COLUMN_WHITELIST = {"schema_common.py", "settlement_rules.py"}
_MONEY_WHITELIST = {"settlement_rules.py"}


def _read(name, directory=_SCRIPTS):
    with open(os.path.join(directory, name), encoding="utf-8") as fh:
        return fh.read()


def _own_nodes(node):
    """作用域内的直接节点；不进入嵌套函数、Lambda 与类（它们各自单独成作用域）。"""
    stack = list(ast.iter_child_nodes(node))
    while stack:
        current = stack.pop()
        if isinstance(current, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda, ast.ClassDef)):
            continue
        yield current
        stack.extend(ast.iter_child_nodes(current))


def _subtree(node):
    stack = [node]
    while stack:
        current = stack.pop()
        yield current
        stack.extend(ast.iter_child_nodes(current))


def _scopes(tree):
    """返回 (作用域名, 参数名, 作用域节点)；模块级代码作为一个无名作用域。"""
    scopes = [("<module>", [], tree)]
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            args = node.args
            names = [a.arg for a in getattr(args, "posonlyargs", []) + args.args + args.kwonlyargs]
            scopes.append((node.name, names, node))
    return scopes


def _hashlib_aliases(tree):
    hashlib_names, sha_direct = set(), set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "hashlib":
                    hashlib_names.add(alias.asname or "hashlib")
        elif isinstance(node, ast.ImportFrom) and node.module == "hashlib":
            for alias in node.names:
                if alias.name in ("sha256", "new"):
                    sha_direct.add(alias.asname or alias.name)
    return hashlib_names, sha_direct


def _is_sha_call(call, hashlib_names, sha_direct):
    func = call.func
    if isinstance(func, ast.Name):
        return func.id in sha_direct
    if not isinstance(func, ast.Attribute):
        return False
    rooted = isinstance(func.value, ast.Name) and func.value.id in hashlib_names
    if not rooted:
        return False
    if func.attr in ("sha256", "file_digest"):
        return True
    if func.attr == "new":
        for arg in call.args:
            if isinstance(arg, ast.Constant) and isinstance(arg.value, str) \
                    and "sha256" in arg.value.lower():
                return True
        for keyword in call.keywords:
            if keyword.arg in ("alg", "digestmod") and isinstance(keyword.value, ast.Constant) \
                    and "sha256" in str(keyword.value.value).lower():
                return True
    return False


def _has_streaming_feed(nodes):
    """流式哈希喂食惯用法：iter(lambda: fh.read(N), b"") 或循环内 digest.update(...)。

    注意不要用「循环内 .read」当签名——zipfile/docx 逐部件读取（verify_fill 的
    package_fingerprint）也是循环内 read，但做的是部件内容摘要，语义合法。
    """
    for node in nodes:
        if isinstance(node, ast.Lambda) and any(
                isinstance(c, ast.Call) and isinstance(c.func, ast.Attribute) and c.func.attr == "read"
                for c in _subtree(node)):
            return True
        if isinstance(node, (ast.While, ast.For)) and any(
                isinstance(c, ast.Call) and isinstance(c.func, ast.Attribute) and c.func.attr == "update"
                for c in _subtree(node)):
            return True
    return False


def _scan(directory=_SCRIPTS):
    """返回 [(文件, 行号, 能力, 说明), ...]；空列表即通过。"""
    findings = []
    for name in sorted(n for n in os.listdir(directory) if n.endswith(".py")):
        tree = ast.parse(_read(name, directory), filename=name)
        hashlib_names, sha_direct = _hashlib_aliases(tree)
        for scope_name, params, node in _scopes(tree):
            own = list(_own_nodes(node))
            if name not in _SHA_WHITELIST:
                sha_calls = [n for n in own
                             if isinstance(n, ast.Call) and _is_sha_call(n, hashlib_names, sha_direct)]
                file_digest_used = any(n.func.attr == "file_digest" for n in sha_calls)
                if sha_calls and (_has_streaming_feed(own) or file_digest_used):
                    findings.append((name, node.lineno, "sha256",
                                     "函数 %s 自建流式文件 SHA-256；唯一实现在 schema_common.sha256_file"
                                     % scope_name))
            if name not in _BARCODE_WHITELIST:
                for n in own:
                    if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute) \
                            and n.func.attr == "endswith" and any(
                                isinstance(a, ast.Constant) and a.value == ".0" for a in n.args):
                        findings.append((name, n.lineno, "barcode",
                                         "出现 \".0\" 剥尾惯用法；条码归一化唯一实现在 "
                                         "schema_common.normalize_barcode"))
                        break
            if name not in _COLUMN_WHITELIST and scope_name != "<module>":
                alias_params = {p for p in params if "alias" in p.lower() or p == "names"}
                header_params = {p for p in params if "header" in p.lower()}
                if alias_params and header_params:
                    for n in own:
                        if isinstance(n, ast.For) and isinstance(n.iter, ast.Name) \
                                and n.iter.id in alias_params \
                                and any(isinstance(c, ast.Compare)
                                        and any(isinstance(op, ast.In) for op in c.ops)
                                        for c in _subtree(n)):
                            findings.append((name, n.lineno, "column",
                                             "函数 %s 对别名参数做表头 in 匹配；列定位唯一实现在 "
                                             "schema_common.locate_column / map_field_columns" % scope_name))
                            break
            if name not in _MONEY_WHITELIST:
                for n in own:
                    if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute) \
                            and n.func.attr == "quantize":
                        findings.append((name, n.lineno, "money",
                                         "Decimal.quantize 出现在 settlement_rules 之外；"
                                         "金额舍入唯一实现在 settlement_rules.money"))
                        break
    return findings


class AstCapabilityScanTest(unittest.TestCase):
    """真实仓库必须零疑似违例。"""

    def test_scripts_directory_is_clean(self):
        findings = _scan()
        self.assertEqual(findings, [],
                         "scripts/ 存在疑似共享内核重实现（capability 文件:行 说明）：\n" +
                         "\n".join("%s:%s [%s] %s" % f for f in findings))


class GateSelfTest(unittest.TestCase):
    """门禁自身的拦截力验证：把复核发现的两类绕过场景原样种进临时目录。"""

    def _scan_planted(self, files):
        with tempfile.TemporaryDirectory() as tmp:
            for name, source in files.items():
                with open(os.path.join(tmp, name), "w", encoding="utf-8") as fh:
                    fh.write(source)
            return _scan(tmp)

    def test_planted_file_sha_via_new_alias_is_caught(self):
        # 绕过场景一（复核 F1）：hashlib.new("sha256") + import 别名 + 打开文件。
        source = (
            "import hashlib as h\n"
            "import json\n"
            "\n"
            "def file_digest(path):\n"
            "    d = h.new('sha256')\n"
            "    with open(path, 'rb') as fh:\n"
            "        for chunk in iter(lambda: fh.read(1024), b''):\n"
            "            d.update(chunk)\n"
            "    return d.hexdigest()\n"
        )
        findings = self._scan_planted({"sneaky.py": source})
        self.assertTrue(any(kind == "sha256" for _, _, kind, _ in findings), findings)

    def test_planted_barcode_normalizer_is_caught(self):
        source = "def my_norm(value):\n    text = str(value).strip()\n    return text[:-2] if text.endswith('.0') else text\n"
        findings = self._scan_planted({"bcode.py": source})
        self.assertTrue(any(kind == "barcode" for _, _, kind, _ in findings), findings)

    def test_planted_alias_column_loop_is_caught(self):
        source = (
            "def find(headers, aliases):\n"
            "    for name in aliases:\n"
            "        for i, header in enumerate(headers):\n"
            "            if name in header:\n"
            "                return i\n"
        )
        findings = self._scan_planted({"cols.py": source})
        self.assertTrue(any(kind == "column" for _, _, kind, _ in findings), findings)

    def test_planted_quantize_outside_authority_is_caught(self):
        source = (
            "from decimal import Decimal, ROUND_HALF_UP\n"
            "def my_money(value):\n"
            "    return Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)\n"
        )
        findings = self._scan_planted({"cash.py": source})
        self.assertTrue(any(kind == "money" for _, _, kind, _ in findings), findings)

    def test_json_digest_without_file_open_is_not_flagged(self):
        # findings_digest / values_digest 类 JSON 摘要是合法语义，不得误伤。
        source = (
            "import hashlib\n"
            "import json\n"
            "def digest(payload):\n"
            "    return hashlib.sha256(json.dumps(payload, sort_keys=True).encode('utf-8')).hexdigest()\n"
        )
        self.assertEqual(self._scan_planted({"ok.py": source}), [])

    def test_in_memory_bytes_digest_is_not_flagged(self):
        # 已读入内存的字节串做一次原语摘要（附件摘要 / locate_confirm 的 raw）是合法语义。
        source = (
            "import hashlib\n"
            "import json\n"
            "def process(path):\n"
            "    with open(path, 'rb') as fh:\n"
            "        raw = fh.read()\n"
            "    json.loads(raw)\n"
            "    return hashlib.sha256(raw).hexdigest()\n"
        )
        self.assertEqual(self._scan_planted({"ok2.py": source}), [])


def _read_scripts(name):
    return _read(name)


class CanonicalHomesTest(unittest.TestCase):
    """四类能力的唯一权威实现必须存在。"""

    def test_kernel_primitives_exist(self):
        body = _read_scripts("schema_common.py")
        for marker in ("def sha256_file", "def normalize_barcode", "def locate_column",
                       "def map_field_columns"):
            self.assertIn(marker, body)

    def test_money_authority_exists(self):
        body = _read_scripts("settlement_rules.py")
        self.assertIn("def money", body)
        self.assertIn("ROUND_HALF_UP", body)


class WiringTest(unittest.TestCase):
    """调用方必须走共享实现（保留早期门禁的有效断言）。"""

    def test_both_engines_use_locate_column(self):
        for name in ("recon_engine.py", "monthly_recon_engine.py"):
            self.assertIn("locate_column", _read_scripts(name))

    def test_monthly_delegates_barcode(self):
        self.assertIn("normalize_barcode", _read_scripts("monthly_recon_engine.py"))

    def test_check_policy_uses_shared_mappers(self):
        body = _read_scripts("check_policy.py")
        self.assertIn("map_field_columns", body)
        self.assertIn("rate_value_column", body)
        self.assertIn("normalize_barcode", body)

    def test_engines_import_shared_sha(self):
        for name in ("recon_engine.py", "monthly_recon_engine.py"):
            self.assertIn("sha256_file", _read_scripts(name))
        # review_common / contract_release / plan_trust 不得再定义流式文件哈希函数；
        # 它们保留的 hashlib 原语调用（findings_digest / values_digest）由 AST 能力扫描管辖。
        for name in ("review_common.py", "contract_release.py", "plan_trust.py"):
            body = _read_scripts(name)
            self.assertNotIn("def sha256_file", body)
            self.assertNotIn("def file_sha256", body)
            self.assertIn("sha256_file", body)

    def test_writeback_and_review_anchor_on_money(self):
        for name in ("writeback_ledger.py", "review_settlement.py"):
            self.assertIn("from settlement_rules import money", _read_scripts(name))

    def test_recon_engine_has_no_first_hit_alias_loop(self):
        self.assertNotIn("for name in names", _read_scripts("recon_engine.py"))


class KernelPinningSuitePresentTest(unittest.TestCase):
    """钉死测试（golden vector）必须随内核存在。"""

    def test_pinning_suite_present(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_shared_kernel.py")
        self.assertTrue(os.path.isfile(path), "tests/test_shared_kernel.py 钉死套件不得删除")


if __name__ == "__main__":
    unittest.main()
