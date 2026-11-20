# PR #1 审计报告（完整版）

> **历史审计记录**：文中 P0/P1/P2/P3 与 N1–N7 均已在后续轮次全部修复（41/41 闭环，见第十–十二节）；本文仅作复盘留痕，不代表当前质量状态。

**仓库**：`KentQi/dealer-ops-suite`
**PR**：[#1 Implement dealer contract and reconciliation workflows](https://github.com/KentQi/dealer-ops-suite/pull/1)
**审计对象**：`feat/dealer-ops-workflow` @ `55172f2`，基线 `main` @ `c486788`
**规模**：37 个文件，+5231 / −2349，1 个 commit
**审计日期**：2026-11-20
**结论**：**不建议合并。** 客户主档建库链路在新版下完全走不通（实测 exit 1、整单不写入），另有 2 条会直接产生错误金额、5 条会让本该阻断的数据静默通过。

---

## 一、审计方法与覆盖

| 项 | 内容 |
|---|---|
| 规模 | `git diff main...HEAD` 全量 9006 行，10 个独立审查角度并行 |
| 角度 | 逐行扫描、删除行为审计、跨文件追踪、Python/openpyxl 陷阱、分层包装、复用、简化、效率、修复层次、规范符合性 |
| 候选 | 72 条；经逐条对抗式验证：**57 CONFIRMED / 7 PLAUSIBLE / 8 REFUTED** |
| 补漏 | 另起一轮找缺口，新增 6 条候选 |
| 实测 | 关键金额与阻断类结论均在 `/tmp` 独立脚本中复现，未改动仓库文件 |
| 测试 | `python -m unittest discover -s tests -v` → **70 passed**，与 PR 声明一致 |
| 已排除 | 18 条曾怀疑但证伪的点（第四节），避免重复排查 |

标注说明：**实测** = 已跑出具体输出；**读码** = 两侧代码逐行对照确认；**文档** = 文件内容直接比对；**待复核** = 专家提出、我未亲验，仅作线索。

---

## 二、P0 · 阻断级（合并前必须修）

### P0-1【实测】客户主档建库链路完全走不通：`apply_contract` 拒绝所有合同字段

**位置**：`apply_contract.py:213-220`；`extract_contract.py:138-235`；`make_confirm_html.py:633`

**事实链**：

1. `extract_contract.py` 为每个合同字段生成 `target="合同.起始"`、`"合同.终止"`、`"合同.结算方式"`、`"合同.付款日"`……，**不带 `@合同编号=` 主键**。
2. `make_confirm_html.py` 提交时原样透传：`fields.push({field, target, value, from})`，同样不补主键。
3. 本 PR 给 `apply_contract.py` 新增了主键协议：

```python
if col == "合同编号" and not keys:
    ...                       # 只有「合同编号」字段本身能追加新行
else:
    if set(keys) != {"合同编号"}:
        raise ValueError("续签／修改合同字段必须带 @合同编号=原合同编号，不默认第一份")
```

**实测**（用 `make_master.py` 新建主档，模拟 `extract_contract` 的真实输出）：

```
exit: 1
拒绝  合同起始     续签／修改合同字段必须带 @合同编号=原合同编号，不默认第一份
拒绝  合同终止     同上
拒绝  结算方式     同上
拒绝  付款日       同上
合计：可写入 0 项，冲突 1 项，跳过 1 项，拒绝 4 项
[中止] 新政策／新合同行存在被拒字段，本次整单不写入，避免留下只有主键的残行。
```

**后果**：新客户建主档时，合同编号能追加，但起止日、结算方式、付款日等**全部落不了库**，且整单回滚。主档缺合同起止 → `load_master.select_contract` 找不到覆盖本期的合同 → 对账链路直接中止。这是把「谈客户 → 建台账 → 对账」的入口堵死。

**修复**：三选一——① `extract_contract` 在合同字段 target 上补 `@合同编号=<已抽取的编号>`；② `make_confirm_html` 提交时用 `_合同编号` 统一为所有 `合同.*` 字段补主键；③ `apply_contract` 在没有 `@` 主键但主档仅有**唯一一行**合同时自动绑定，多行时才阻断。推荐 ①+③ 组合（抽取端带主键、写入端保留单行兜底）。

---

### P0-2【读码】01 对账单「合计」与「人民币（大写）」自相矛盾

**位置**：`recon_engine.py:1903`（合计）vs `:1908`（大写）

```python
c = ws.cell(row=r, column=7); num(c, T["settle"], "#,##0.00")          # 合计：不含盈亏
...
mput(ws, r, 4, r, 7, rmb_upper(SETTLEMENT_TOTAL), ...)                 # 大写：含盈亏
```

`SETTLEMENT_TOTAL = RULE_RESULT["settlement_amount"]`，在 `VARIANCE_INCLUDED` 时等于 `T["settle"] + 盈亏`。旧版大写用的是 `T["settle"]`，本 PR 改成了 `SETTLEMENT_TOTAL` 但**没有同步改合计行**。

**触发条件**：本期存在盈亏且获得单次授权纳入结算总额。

**后果**：一张要签章的对账单上，紧邻的两行数字对不上（如合计 100,000.00，大写「壹拾万零伍仟元整」）。这是对外单据，双方核对时必然被质疑。

**修复**：明确「合计」行的语义——若它是本期结算总额，改用 `SETTLEMENT_TOTAL`；若它是流水结算金额，大方写应回退为 `T["settle"]`，并在标签上区分「流水合计」与「结算总额」。

---

## 三、P1 · 高（会产生错误金额或让阻断失效）

### P1-1【实测】两条对账链路从同一主档取出不同合同价

**位置**：`recon_engine.py:1053` / `monthly_recon_engine.py:278`

```python
# 铺货
PRICE_BOOK = _load_rate_book(_rate_wb, "价格", ("供货价", "结算价", "合同价"), "合同价", True)
# 月结
prices = load_rate_book(master, "价格", ("结算价", "合同价", "供货价"), True)
```

铺货的 `col()` 把别名循环放外层、第一个命中即返回，且不检查多列命中；月结先精确匹配、同级多列命中即阻断。**实测**同一主档（同时有 `供货价` 与 `结算价` 两列）：

```
recon_engine        取到的合同价 = 28
monthly_recon_engine 取到的合同价 = 33
```

**修复**：取价列解析收敛为一份共享实现（精确匹配优先、多列命中阻断），并显式规定「价格表不得同时存在两个价格列」。

---

### P1-2【实测】政策自检把商品名称列绑到条码列，缺条码阻断永不触发

**位置**：`check_policy.py:28`（`COL_ALIASES["name"]` 含裸 `"商品"`）、`:51-66`、`:84-87`

`"商品"` 是 `"商品条码"` 的子串，`map_columns` 逐列扫描时第 1 列同时命中 `barcode` 与 `name`；第 2 列「商品名称」永不被映射。随后 `collect_rows` 对「条码与品名都为空」的行执行 `continue`（注释：表尾说明）。

**实测**：一张第 3 行条码留空、品名与价格齐全的政策表：

```
mapping = {'barcode': 1, 'name': 1, ...}
analyse() → blocking_findings: [], summary.blocker: 0
```

**后果**：`P3 条码为空` 这条 blocker 永远不触发；合同包审阅与最终放行都会认为政策表干净。同时审阅报告以条码代替品名。

**修复**：`name` 别名去掉裸 `"商品"`；`map_columns` 增加「一列不得映射两字段」守卫；跳行条件改为「整行有效单元格为空」。

---

### P1-3【读码】非结算类甲方行被强制校验合同价，正常对账直接中止

**位置**：`recon_engine.py:1350-1351`

```python
for _bar in sorted(set(opening_by_bar) | {line["barcode"] for line in _rule_sales}):
    PRICE_BOOK.validate_period(_bar, P_START, P_END)
```

`opening_by_bar` 来自全部甲方挂账行，**未按结算属性过滤**。若某个「非结算类」SKU（陈列、赠品、试用装）没有连续合同价，`validate_period` 抛 `RATE_COVERAGE_GAP`，整次对账中止。而文档明确说非结算类只需登记并被排除，不需要结算价。

**修复**：校验前按 `MS_BAR[bar]["attr"] == "结算类"` 过滤；非结算类只校验条码有效性。

---

### P1-4【读码】非结算类 SKU 进入数量勾稽，产生虚假盈亏

**位置**：`recon_engine.py:1352-1359`（传入未过滤的 `opening_by_bar` / `inventory_by_bar`）

`_rule_sales` 已按 `attr == "结算类"` 过滤，但 `calculate_consignment` 的 `opening` / `inventory` 是全量。非结算类 SKU 没有销售流水，`theoretical_end = opening`，`adjusted = inventory`，一旦两者不等就产出「短缺／盘盈·待双方核对」，并计入 `variance_display_amount`。文档明确「非结算类商品必须明确登记并被排除」。

**修复**：构造 `opening_by_bar` / `inventory_by_bar` 时即按结算属性过滤；或传入后由规则层按属性剔除并单独列示「未参与勾稽」。

---

### P1-5【读码】销售流水缺流水号或品名的行被静默跳过

**位置**：`recon_engine.py:966-967`

```python
if not r or not _cell_of(r, o_id) or not _cell_of(r, o_name):
    continue
```

缺流水号或缺商品名的销售／退货行被直接丢弃，**无告警、无计数、无阻断**。结算数量、金额、理论结存随之少算，且对账结果看起来一切正常。圆括号里的「流水号」正是退货追溯的原键——丢了它，退货也无法追溯。

**修复**：缺流水号或缺品名 → 阻断并列出源表行号；至少写入 `warnings` 并在对账单上标注「已排除 N 行」。

---

### P1-6【读码】铺货回写载荷缺合同编号与业务类型，且凭证校验拦不住

**位置**：`recon_engine.py:2972-2980`；`writeback_ledger.py:308-314`

月结回写写了 `contract_no` / `business_type`，铺货没写；而 `CONTRACT_NO` 就在第 131 行、同一作用域内。更糟的是凭证校验：

```python
def _same_writeback_value(key, left, right):
    ...
    return left == right          # None == None → True
```

两侧都缺失时判「一致」，所以这道闸门无法发现载荷本身缺字段。

**后果**：多合同客户回写后，「期间」行的合同编号与业务类型留空，台账无法归属到任何合同；后续按合同汇总、到期提醒、复核都失去锚点。

**修复**：载荷补两个字段；关键字段（`contract_no`、`business_type`、`settle_amount`）要求两侧均存在才判一致。

---

### P1-7【实测】待办 id 撞车

**位置**：`make_todos.py:127-128`

`raw_id = "|".join((合同编号, 期间, 类型, 标题, 截止日))`，而履约提醒的 `title` 只含条款类型。

**实测**：同一合同两条返利条款（季度 3% / 年度 5%）→ 同一个 `TODO-bd7a3410c785`，6 条待办只有 5 个唯一 id。下游按 id 去重或更新状态时，一条会被覆盖。

**修复**：`raw_id` 纳入 `key_value` 与来源位置；或对重复 id 直接报错。

---

### P1-8【读码】复核包强制依赖 `RECON_CFG`，SKILL.md 的复核命令在新会话必失败

**位置**：`make_review_packet.py:1210-1212`（CLI 无 `--cfg`）；`_read_cfg()` 只读环境变量；`review_settlement.py:63-64`

SKILL.md 第 5 节的命令是：

```bash
python3 make_review_packet.py "<对账工作簿.xlsx>" --master "<_主档.xlsx>" --out ... --md ...
```

没有 `RECON_CFG`。而 `_read_cfg` 在环境变量缺失时返回 `None` → `sources` 为空 → `validate_settlement_inputs` 抛「复核包无法读取源文件（role=jia），不能沿用原结算结果」。每个 bash 块通常在新 shell 执行，第 4.2 节的 `export` 不会带过来。

**修复**：给 `make_review_packet.py` 增加 `--cfg` 参数（文档同步），或在 SKILL.md 第 5 节的命令里显式 `export RECON_CFG="$(cat ...)"`。

---

### P1-9【文档】客户主档人工确认门槛被整体移除

**位置**：`SKILL.md:128-129`；`references/客户主档设计.md`

旧版有 A1–A5 完整流程与两条红线：「A3 确认门槛不可跳过：抽取结果是候选，不是结论」「A3.5 不得要求用户移动确认单」。新版 SKILL.md 只剩两条 `apply_contract.py <候选字段.json>` 命令；`客户主档设计.md` 中「确认页／确认单／逐项确认／候选字段／抽取」**零命中**；`extract_contract.py` / `locate_confirm.py` / `make_confirm_html.py` / `make_digest.py` 在 SKILL.md 与全部 references 中都不再被提及。

`apply_contract.py` 的确认人校验只在输入是确认单时生效：

```python
confirmed = data.get("schema") == "contract-confirm/v1"
if confirmed and not str(meta.get("operator") or "").strip(): ...
```

普通候选 JSON 加 `--write` 照常落盘。

**后果**：Agent 照新版 SKILL.md 执行，可以跳过人工逐项确认直接写主档——违背项目「不替用户做业务决定」原则。

**修复**：恢复确认流文档；或让 `apply_contract` 对非确认单的 `--write` 强制要求 `--confirmed-by`。

---

## 四、P2 · 中

| # | 位置 | 问题 | 证据 |
|---|---|---|---|
| P2-1 | `make_todos.py:150-161` | 付款日写成「月结30天」「账期30日」时，铺货既不生成付款待办也无任何告警；月结才有「资料缺失」兜底。业务类型判断用原始字符串 `== "月结买断"`，写「月结」时兜底也不触发 | 实测 |
| P2-2 | `recon_engine.py:1381` | 差异数量 `int()` 截断而金额保留到分。实测实数 9.5 → 台账「差异数量 0、差异金额 −5000.00」 | 实测 |
| P2-3 | `make_todos.py:81` vs `ledger_read.py:77` | 两套工作日算法（节假日日历 vs 只排周末）。实测同一 5 个工作日：10-08 vs 10-13 | 实测 |
| P2-4 | `make_todos.py:28` | 斜杠日期被静默当作缺失。实测 `2027/01/01` 的合同不被过滤、`2026/12/31` 的终止日不生成到期提醒，warnings 全空 | 实测 |
| P2-5 | `monthly_recon_engine.py:168/281` | 月结不认 `_数据映射.json` 的 `master.sheets` 改名（铺货认），同一主档一条链路能跑一条报「缺少价格表」 | 读码 |
| P2-6 | `settlement_rules.py:250` | 销售贴花取期末标准、退货按原销售日标准冲减；期中调标时，同期的销售+完整退货会留下贴花残值 | 读码 |
| P2-7 | `review_contract_package.py:319/449-457` | 政策 blocker 只在 `--release/--approve-write` 时升级；普通审阅里降级为 high，`summary.blocker=0`、退出码 0 | 读码 |
| P2-8 | `writeback_ledger.py:333` | 两个引擎的 `_回写.json` 都不含 `statement_no`；凭证校验拒绝本次 CLI 传值，所以对账单编号永远写不进「期间」表 | 读码 |
| P2-9 | `make_todos.py:99-105` | 付款顺延跨年时只用 `due.year` 判断日历覆盖，顺延后落入未覆盖年份仍标 `estimated=False`，给出伪精确日期 | 读码 |
| P2-10 | `settlement_rules.py:226` | 引擎总传 `original_decal_charged=None`，规则层回查历史 JSON 的兜底分支不可达；退货贴花依据实际只认退货行自身字段 | 读码 |
| P2-11 | `apply_fill.py:228` | 缺省 `--spec/--values` 时静默跳过输入绑定核对，写前门禁可被绕过 | 待复核 |
| P2-12 | `contract_release.py:154` | 放行凭证未绑定写前审阅的 values 输入，理论上可用未复核的 values 放行 | 待复核 |
| P2-13 | `make_review_packet.py:836` | 映射显式指定的源文件与复核包重算路径不一致时，带映射的账无法复核 | 待复核 |

---

## 五、P3 · 低 / 文档与规范

| # | 位置 | 问题 |
|---|---|---|
| P3-1 | `SKILL.md:97` | `verify_fill.py --plan` 的产出脚本 `build_plan.py` 未再路由，只出现在 `references/合同编辑.md` |
| P3-2 | `SKILL.md:161` | 签章 PDF 链路（`load_master --pdf` → `apply_v17`）在指令层断开；frontmatter 与 README 仍宣称支持 |
| P3-3 | `README.md:27` | 仍称政策自检「只做三类」，实际已五类（P1/P3/P4/P5/P6） |
| P3-4 | `SKILL.md:1` | 净删约 669 行未按项目规范归档旧版（规范要求移入 `archive/`） |
| P3-5 | `references/结算规则.md` | 与 SKILL.md／复核协议／客户主档设计重复同一套口径，形成多份权威（违反「单一权威载体」） |
| P3-6 | `docs/source/architecture.json`、`business-workflow.json` | 架构图/流程图未随 PR 更新（脚本数、文档数、待办入口均为旧值；流程图仍称每月 1 日自动提醒，与新规格「不创建定时任务」冲突） |
| P3-7 | `review_contract_package.py:327` | `merge_llm` 对顶层数组输入抛 `AttributeError`（`data.get` 先于 `isinstance` 求值），而错误文案承诺支持数组 |
| P3-8 | `load_master.py:127` | 重写后丢掉年份范围校验，`000001` 会抛原始 traceback 而非中文报错 |
| P3-9 | `tests/test_monthly_recon_end_to_end.py:16` | 不能按文件直接运行（`from tests.` 但 `tests` 不是包） |
| P3-10 | `recon_engine.py:1401` | 重写后遗留旧聚合变量（13 个死变量 + 与 `RULE_RESULT` 重复的 `T` 总计） |

---

## 六、结构性改进（不是单点 bug，是复发根因）

| 主题 | 现状 | 建议 |
|---|---|---|
| 列定位 | 至少三套实现（铺货 `col()`、月结 `column()`、政策 `map_columns()`），严格度与别名顺序各不相同 | 收敛到 `schema_common.py` 一份，规则：精确优先、多列阻断、同列不得映射两字段 |
| 条码归一化 | 三套（`settlement_rules.barcode_of`、`check_policy.norm_barcode`、月结 `bar`） | 同上 |
| 金额舍入 | 四套（规则层 `money`、引擎 `round`、复核 `_number`、回写 `parse_number`），`.xx5` 输入下合计与明细可能不等 | 统一 `ROUND_HALF_UP` + 逐行舍入后合计 |
| SHA-256 | 四份副本（`review_common`、`contract_release`、`recon_engine`、`monthly_recon_engine`） | 单一实现，避免门禁各算各的 |
| 复核包体积 | 同一批逐笔流水内联三份（`rule_result` + `priced_rows` + `tiehua_rows`） | 只保留投影后的一份，`rule_result` 按需引用 |
| 主档解析 | `recon_engine` 相邻 14 行内把同一主档完整解析两次 | 复用已打开的工作簿 |

---

## 七、已排除的怀疑（不必重复排查）

| 曾怀疑 | 核对结果 |
|---|---|
| 月结 `opening_qty/diff_qty = 0` 违反「必须为 null」 | 协议针对 `figures`；复核包按标签扫描工作簿，无对应标签时置 `null` 并告警 |
| `contract_release.py`、`review_settlement.py` 是死代码 | 分别被 `apply_fill`、`review_contract_package`、`verify_fill`、`make_review_packet` import |
| 放行闸门把政策 blocker 降级 | `--release/--approve-write` 路径显式保留并升级（普通审阅的问题见 P2-7） |
| `evidence` 是字符串导致等级升级失效 | `add()` 统一存 `{"value": ...}`；字符串路径由 `JSONDecodeError`（ValueError 子类）捕获 |
| 模板「期间」表 17 列与回写 FIELDS 错位 | 逐列比对完全一致 |
| 条码 8–14 位规则与项目数据不符 | 测试与模板均为 13 位数字 |
| 铺货期间闭区间导致期末当天漏算 | 业务上期末当天应计入，正确 |
| `cfg["period"]["next_month"]` 存整数是类型错误 | 下游按 `int` 消费 |
| `apply_v17` 保留旧 `rule_result` 是漏更新 | v17 只加附表不改金额，设计如此 |
| 测试污染仓库根目录 | 跑完即清理，实测为 0 |
| 政策表 `blocking_findings` 被审阅端丢弃 | `analyse()` 把全部发现都放进 `needs_confirmation` |
| `review_common` / `review_settlement` 路径校验可绕过 | 哈希与 realpath 双重校验，链路完整 |
| 月结结果凭证缺主档绑定 | `master`/`workbook`/`sources` 哈希齐全 |
| 合同包双门禁可绕过 | `apply_fill` 要 `--review`、`verify_fill` 要 `--review-release`，绑定五类输入 SHA-256 |
| 复核结论可带失败项标「通过」 | `validate_report` 在三类条件下分别禁止 |
| 铺货回写凭证可复用旧结果 | 主档 + 工作簿 + 锁 pre/post 哈希 + 归档报告四重绑定 |
| 政策表检查根本没被调用 | `review_contract_package.check_policy` 子进程调用，退出码 0/2 都处理 |
| 条码映射在 `_数据映射.json` 里方向反了 | `recon_engine` 正确反转 `{actual: canonical}` |

---

## 八、修复优先级

**合并前必须修**

1. P0-1 客户主档建库断链（阻断整个入口）
2. P0-2 01 表大写与合计矛盾（对外单据自相矛盾）
3. P1-1 / P1-3 / P1-4 取价与结算属性过滤（直接错钱）
4. P1-2 政策自检漏阻断（门禁形同虚设）
5. P1-5 销售行静默丢弃（错钱且无痕）
6. P1-6 + P2-8 回写字段与凭证校验
7. P1-7 待办 id
8. P1-8 复核包 CFG 依赖（复核步骤跑不通）
9. P1-9 确认门槛（需要拍板：恢复文档还是脚本强制）

**建议同批修**

P2-1 付款义务、P2-2 截断、P2-3 工作日、P2-4 斜杠日期、P2-5 表名映射、P2-6 贴花残值、P2-7 blocker 等级。

**随后跟进**：P2-9~13、P3 全部、第六节结构性收敛（能预防同类缺陷复发）。

---

## 九、覆盖声明

- 10 个审查角度全部完成，72 条候选全部经过对抗式验证，另有 6 条补漏候选。
- P0-1、P0-2、P1-1、P1-2、P1-3/4、P1-5、P1-6、P1-7、P1-8、P1-9 由我逐条复核（代码对照或实测）。
- P2-11~13 与部分 P3 来自专家角度，标注「待复核」，未计入合并阻断理由。
- 本报告只做只读审查，仓库代码未改动；唯一新增文件是本报告与临时探针（均在 `/tmp`）。

---

## 十、修复复核（2026-11-20，工作区未提交改动）

**结论：11 条 P0/P1 全部落地，10 条确认修好；其中 2 条在修复过程中引入了新回归，需要再收一次尾。**

### 10.1 逐条核对

| 原编号 | 状态 | 复核证据 |
|---|---|---|
| P0-1 建库断链 | ✅ | **实测全链路**：`extract_contract` → `make_confirm_html` → `apply_contract`。抽取端自动带主键（`合同.起始@合同编号=HT-2026-009`），确认页按最终编号重建 target，写入落在**新合同行**，模板示例行 `HT-DEMO-2026-001` 未被覆盖；exit 0、拒绝 0 项 |
| P0-2 合计 vs 大写 | ✅ | **实测**：盈亏授权纳入时，合计行 = 3265，大写 = 叁仟贰佰陆拾伍元整；行标签改为「本期结算总额」，第二节③④⑤ 说明同步 |
| P1-1 取价列分叉 | ✅ | 新增 `settlement_rules.rate_value_column()`，两引擎共用；多列命中 → `AMBIGUOUS_RATE_COLUMN` 阻断；别名顺序差异已消除（多列即阻断，顺序不再影响结果） |
| P1-2 政策自检漏阻断 | ⚠️ | 目标达成（裸「商品」别名已删、同列双映射守卫已加、缺条码行已阻断），但**引入 N1** |
| P1-3 非结算类强制合同价 | ✅ | 非结算类在 `jia` 循环即剔除，不进 `opening_by_bar` → 不触发 `validate_period`；实测 exit 0 |
| P1-4 非结算类进入勾稽 | ✅ | **实测**：甲方 500 件 + 库存 999 件非结算 SKU 下，`opening_qty=100`（未被污染）、`diff_qty=3` 不变、`settle=3100` |
| P1-5 销售行静默丢弃 | ⚠️ | 目标达成（缺流水号/品名 → 阻断并给行号），但**引入 N2** |
| P1-6 回写载荷缺字段 | ✅ | `_回写.json` 补 `contract_no`/`business_type`/`statement_no`；`verify_write_provenance` 强制三者非空，`None == None` 漏洞闭合 |
| P1-7 待办 id 撞车 | ✅ | `raw_id` 纳入 source/key_value/evidence；新测试断言全部 id 唯一 |
| P1-8 复核包依赖 `RECON_CFG` | ✅ | 新增 `--cfg`；新测试在**删掉环境变量**后仍通过；SKILL.md 与 `复核协议.md` 已同步 |
| P1-9 确认门槛被移除 | ✅ | 普通候选 JSON `--write` 强制 `--confirmed-by`；SKILL.md 与 `客户主档设计.md` 恢复 `extract → 确认页 → locate_confirm → apply` 全链路 |

**顺带修好的 P2-8**：`statement_no` 全链路打通（`load_master --statement-no` → `RECON_CFG` → `_回写.json` → 台账「对账单编号」列，该列已在 `writeback_ledger.FIELDS` 中）。

**测试**：`python -m unittest discover -s tests` → **84 passed**（原 70 + 新增 14）。测试残留目录实测 0（第四节那条「测试污染仓库根目录」仍属证伪）。

### 10.2 新引入的回归（本轮必须再收一次尾）

#### N1【实测 · P1】政策表的表尾说明行被误判为阻断

**位置**：`check_policy.py:85-86`（`collect_rows` 跳行条件）

跳行条件从「条码与品名同时为空即跳过」改成「整行有效单元格均为空才跳过」。于是表尾说明行只要**任一列**有字（备注列、依据文件列都算），就成了数据行。

**实测**（政策表 2 行 SKU + 1 行表尾）：

```
data_rows  = 3
blocking   = [('P3', '条码为空', 4), ('P4', '结算价为空', 4), ('P4', '生效日为空', 4)]
```

表尾文字改放「依据文件」列，同样被拦（data_rows=2，blocking 行号 3）。

**旧版是明确跳过的**（`git show HEAD:.../check_policy.py`）：

```python
if not barcode and not name:
    continue  # 表尾说明，不是 SKU 行
```

**影响链**：`check_policy` exit 2 → `review_contract_package` 把每条 `needs_confirmation` 记为「政策表 high」→ 放行路径升级为 blocker → **合同包无法放行**。讽刺的是 `check_policy.collect_footer()` 本来就是专门去扫表尾「起订量/累计」的，修复与该功能自相矛盾。

**修复建议**：把「数据行」判据收窄为 `barcode / name / price / decal / start` 任一非空；只有备注列或依据文件列有字的行仍按表尾跳过。既保住 P1-2 的目标（缺条码但品名+价格齐全 → 阻断），又恢复表尾豁免。

#### N2【实测 · P1】销售流水表带合计行 → 整期无法出账

**位置**：`recon_engine.py:971`、`recon_engine.py:973`

**实测**（正常两笔流水 + 一行合计）：

```
B1  合计行（流水号／品名为空，数量 60） → exit 1
    [自检未通过] 销售流水第 4 行缺少流水号；有效业务行不得静默跳过，请补齐后重新对账
B1b 合计行（品名='合计'，无流水号）   → exit 1
```

**不对称**：挂账表与库存表都有显式豁免（`recon_engine.py:756`：`if nm in ("合计", "小计", "总计", "合 计"): continue`），销售流水表没有。引擎自己的提示语也提到「源表可能有合计行」（`recon_engine.py:1607`）。

**修复建议**：在流水号检查之前先豁免聚合行——本行**任一单元格**文本 ∈ {合计, 小计, 总计, 合 计} 即跳过并计入 warning 计数；其余缺流水号/品名的行仍然阻断。

#### N3【实测 · P2】政策自检与两个引擎对「价格列歧义」结论相反

**位置**：`check_policy.py:29`（`COL_ALIASES["price"]`）、`check_policy.py:51-66`（`map_columns`）

同一张表（同时有 `供货价` 与 `结算价` 两列）：

```
check_policy   → columns={'price': 3(供货价)}，findings = []      # 判定「无问题」
rate_value_column → 阻断：AMBIGUOUS_RATE_COLUMN
```

即**审阅报告说政策表干净，两个对账引擎却拒绝出账**。`map_columns` 对同一字段的多列命中是「先到先得、静默忽略」，而 P1-1 刚把引擎侧收紧为「多列即阻断」，两侧口径分叉。

**修复建议**：`check_policy` 的 `price`／`decal` 列解析改用 `settlement_rules.rate_value_column`，或至少在多列命中时报 blocker（不要静默取左列）。

### 10.3 复核建议

1. 先收 N1、N2 两条回归（各补一条回归测试：表尾行不阻断；合计行不阻断）。
2. 再收 N3（政策自检与引擎共用取价列解析）。
3. P2 段其余 12 条（P2-1~7、P2-9~13）与 P3 十条本轮未动，仍按原报告优先级跟进。

---

## 十一、修复复核（第二轮：P2／P3 段，工作区未提交改动）

**结论：P2 13 条 + P3 10 条 + 上一轮 N1／N2／N3 共 26 项，24 项确认修好，2 项未完全落地。**

改动规模从 19 文件 / +422 行扩大到 **30 文件 / +1175 行**，新增 `scripts/workdays.py` 与 `skills/dealer-ops-suite/archive/`。测试 **70 → 84 → 105 passed**（本轮 +21 条），测试残留目录实测仍为 0。

### 11.1 逐条核对

| 原编号 | 状态 | 复核证据 |
|---|---|---|
| P2-1 付款义务兜底 | ✅ | 付款日无法解析出明确日 → 告警 + 生成「资料缺失」待办；业务类型改用 `"月结" in business_type_of(contract)` 归一化，「月结30天」「日＞28」均不再推定默认日期 |
| **P2-2 数量截断** | ⚠️ **未落地** | 输出层 `_qty()` 已改（不再 `int()`），但**输入层 7 处仍 `int(round(...))`**，实测多算 15.00 元 —— 见 11.2 N4 |
| P2-3 两套工作日算法 | ✅ | 新增 `scripts/workdays.py` 为单一权威实现，`make_todos` 与 `ledger_read` 共用；两者输出均带 `workday_basis` 口径标注，不再混用 |
| P2-4 斜杠日期 | ✅ | `workdays.parse_day` 支持 `%Y/%m/%d`；`optional_day` 收集告警而非静默当作缺失；终止日斜杠日期也能生成到期提醒 |
| P2-5 月结不认表名映射 | ✅ | 月结新增 `master_sheet_name()`，与铺货 `_master_sheet` 同一口径；实测把 价格／贴花／SKU 改名后月结仍正常出账（buyout 60 / 3100.00） |
| P2-6 贴花残值 | ✅ | **实测**：期中调标（10-10 销售 40 件 + 10-25 完整退货 40 件）→ 销售贴花 120.00、退货冲减 −120.00、**净贴花 0.00**，残值消除；本期原销售免贴花标记也能正确冲减；跨期退货缺依据仍按预期阻断 |
| P2-7 blocker 等级降级 | ✅ | `check_policy` 保留自身 severity，放行路径那段「降级再升级」的补丁已删除，普通审阅即 exit 2 |
| P2-8 对账单编号写不进 | ✅ | `statement_no` 全链路打通：`load_master --statement-no` → `RECON_CFG` → `_回写.json` → 台账「对账单编号」列（该列已在 `writeback_ledger.FIELDS` 中） |
| P2-9 付款顺延跨年伪精确 | ✅ | `adjust_payment` 调整后校验日历覆盖年份，超出即返回 `(None, True)` 估算标记，不给精确日期 |
| P2-10 规则层兜底分支不可达 | ✅ | 新增 `_first_present`，显式 `None` 不再挡住历史 JSON 兜底；实测跨期退货缺依据仍阻断 |
| P2-11 写前门禁可绕过 | ✅ | `--spec/--values` 改 `required=True`；增量计划新增 `check_incremental_plan()`，逐条核对字段在规格内、有用户确认值、落笔文本与确认值一致，否则整单不写 |
| P2-12 放行未绑定 values | ✅ | `build_release` 无 `inputs` 直接抛错；`verify_release` 恒要求 `input_files` |
| P2-13 映射源文件与复核路径不一致 | ✅ | `make_review_packet` 优先采用结算结果凭证里的「实际解析后绝对路径 + 哈希」，旧结果才回退 `cfg.base_dir + cfg.sources` 重算 |
| P3-1 `build_plan.py` 未路由 | ✅ | SKILL.md §3 步骤 5 写明编译命令，并注明 plan.json 只能由脚本生成、失败只留 `.rejected` |
| P3-2 签章 PDF 链路断开 | ✅ | SKILL.md 写明 `--pdf` / `--pdf-sheets`；`load_master` 确有这两个参数，`apply_v17` 按 `cfg["pdf"]` 调 `export_pdf`，无 reportlab 时明确跳过并注明，不静默当作已出 |
| P3-3 README「三类」 | ✅ | README 改为五类并列出 P1–P6 六个码 |
| **P3-4 旧版未归档** | ⚠️ **部分修复** | `archive/SKILL-旧版-20261119.md` 已建立，但归档的是**初始发布版**而非本 PR 替换掉的那一版 —— 见 11.2 N5 |
| P3-5 多份权威载体 | ✅ | `references/结算规则.md` 整套口径并入 SKILL.md §4「统一结算口径」，该文件改为指针，不再单独维护口径 |
| P3-6 架构图／流程图未更新 | ✅ | `docs/source/*.json` 已更新（本地待办取代每月 1 日提醒、删除「每月触发」连线、SKILL.md 268 行、references 12 份）；**实测 SKILL.md 恰为 268 行、references 恰为 12 份，图中数值与实际一致** |
| P3-7 `merge_llm` 数组输入 | ✅ | **实测**：顶层数组 → 正常；`dict.findings` → 正常；空 dict → 0 条；字符串 → 明确拒绝，无 `AttributeError` |
| P3-8 `load_master` 年份范围 | ✅ | `month_bounds` 校验 2000–2100 与月份 1–12，非法值给中文报错而非原始 traceback |
| P3-9 测试不能直接运行 | ✅ | **实测** `python3 tests/test_monthly_recon_end_to_end.py` → Ran 10 tests, OK |
| P3-10 死变量 | ✅ | 13 个旧聚合变量（`NOV_VAL`/`ACT_VAL`/`ADJ_VAL`/`THEO_VAL`/`N_PAY`/`OPEN_VAL_P`/`SETTLE_P`/`THEO_VAL_P`/`NOV_VAL_P`/`_absent_end`/`PAY_CN`/`CT_BASIS_TXT` 等）已删，引擎无引用残留 |
| N1 政策表表尾行误阻断 | ✅ | **实测**：备注列表尾 → data_rows=2 / blocking=[]；依据文件列表尾 → data_rows=1 / blocking=[]；缺条码的**数据行**仍阻断（P3 条码为空） |
| N2 销售流水合计行硬阻断 | ✅ | **实测**：品名「合计」「合 计」的聚合行 → exit 0 并打印豁免行数；其余缺流水号／品名的业务行仍阻断并给行号 |
| N3 政策自检与引擎口径相反 | ✅ | **实测**：同一张双价格列表 → 政策自检报 P4「价格／标准列歧义」blocker，与引擎 `AMBIGUOUS_RATE_COLUMN` 同口径 |

### 11.2 仍未完全落地的两项

#### N4【实测 · P1】P2-2 只改了输出层，铺货输入端仍静默取整 → 结算金额算错

**位置**：`recon_engine.py:760`（挂账未结算数量）、`:761`（初始数量）、`:846`（库存数量）、`:911`（次月回加）、`:978`（销售数量）、`:1247`（批次表数量）、`:1280`（促销使用数量）

这 7 处仍是 `int(round(_num_val(...)))`。输出层的 `int(_rule[...])` 虽然换成了不截断的 `_qty()`，但**源表里的实数在进入规则层之前就已经被四舍五入成整数**——对铺货链路来说，P2-2 的修复等于没生效。

**实测**（源表：期初 100 件｜销售 4.7 件｜库存 95.3 件｜合同价 50 元）：

```
exit = 0
引擎结果：期初=100 净动销=5 理论=95 实际=95 盈亏=0
结算金额 = 250.00 元（台账 settle_amount=250.0）
正确应为：销售 4.7 × 50 = 235.00 元｜理论 95.3｜实际 95.3｜盈亏 0
差额 = 15.00 元
```

**为什么查不出来**：对账单上「期初 100 ＝ 净动销 5 ＋ 理论结存 95」照样成立，数量勾稽照样通过，盈亏显示 0，exit 0，无任何告警——错得无声无息。

月结链路没有这个问题（`monthly_recon_engine` 用 `float(line["quantity"])`，不取整）；规则层也是 Decimal 贯穿。**只有铺货链路受影响**。

**修复建议**：

1. 这 7 处改为保留小数的解析（与 `_qty()` 一致，或直接用 Decimal），让规则层拿到原始数量；
2. 若业务上数量必须为整数，则在解析处**阻断**并提示「数量出现小数，请确认源表口径」，而不是静默四舍五入后按错误数量算钱；
3. 补一条端到端回归测试（源表小数 → 结算金额按小数计算，或明确阻断）。

#### N5【实测 · P3】P3-4 归档的是初始发布版，不是本 PR 替换掉的那一版

**位置**：`skills/dealer-ops-suite/archive/SKILL-旧版-20261119.md`

**实测**：

```
归档文件 == cb77c54（初始发布版，843 行）
main(c486788) 中「交付后必须回写台账 / writeback_ledger.py」匹配行数: 6
归档中同类内容匹配行数: 0
```

本 PR 是相对 `main`（c486788，894 行）重写了 SKILL.md，净删 669 行。归档目录里放的是**更早的初始发布版**（cb77c54，843 行），因此 main 相对初始版新增、又被本 PR 删掉的那 51 行——包括「6. 交付后必须回写台账（B5）」这条红线、`writeback_ledger.py` 的完整命令块、「期间表只由 writeback_ledger.py 在 B5 写入」的说明——**至今仍无处可查**。

**修复建议**：把归档件换成 `git show c486788:skills/dealer-ops-suite/SKILL.md` 的内容（或补一份 `SKILL-主分支版-20261119.md`），确保「本 PR 删掉的那一版」原样可查。

### 11.3 整体状态

| 段 | 条数 | 已修好 | 未完全落地 |
|---|---|---|---|
| P0 | 2 | 2 | 0 |
| P1 | 9 | 9 | 0 |
| P2 | 13 | 12 | 1（N4 = P2-2） |
| P3 | 10 | 9 | 1（N5 = P3-4） |
| 复核新增 | 3 | 3 | 0 |
| **合计** | **37** | **35** | **2** |

剩余动作：修 N4（端到端小数数量，否则铺货链路会静默算错钱）与 N5（换正确版本的归档件）。除此之外，原报告第二节「结构性改进」一节的六条（列定位／条码归一化／金额舍入／SHA-256／复核包体积／主档解析）仍未收敛——N4 正是「金额与数量精度口径未统一」这条结构性问题的又一次复现。

---

## 十二、修复复核（第三轮：N1–N5 收尾）

**结论：N1–N7 全部闭环，40 项复核通过，测试 106 项全绿。**

本轮改动 30 文件 / +1220 行 / −331 行，新增回归测试 `test_fractional_quantities_are_not_truncated`。`recon_engine.py` 在复核期间还有一次提交，本节结论以最终工作区为准。

> **审计过程留痕**：第三轮首轮探针跑出「N5 归档件仍是初始版」「N6 测试红」两个结果，但那是**改动落地前的中间状态**——归档件替换与断言修正发生在探针之后。以最终工作区重测，三条均已闭环。下表记录的是最终实测值。

### 12.1 逐条核对（最终状态）

| 编号 | 状态 | 复核证据 |
|---|---|---|
| N1 政策表表尾行误阻断 | ✅ | **实测**：备注列表尾 → data_rows=2 / blocking=[]；依据文件列表尾 → data_rows=1 / blocking=[]；缺条码**数据行**仍报 P3 阻断 |
| N2 销售流水合计行硬阻断 | ✅ | **实测**：品名「合计」「合 计」→ exit 0 并打印豁免行数；其余缺流水号／品名的业务行仍阻断并给行号；无任何标签的裸行仍阻断（无标识行带数量，阻断是合理的） |
| N3 政策自检与引擎口径相反 | ✅ | **实测**：同一张双价格列表 → 政策自检报 P4「价格／标准列歧义」blocker，与引擎 `AMBIGUOUS_RATE_COLUMN` 同口径 |
| N4 铺货输入端静默取整 | ✅ | 7 处取整点全部改为 `_qval()`／`round(..., 4)`（挂账未结算数量、初始数量、库存数量、次月回加、销售数量、批次表数量、促销使用数量）。**实测**：期初 100｜销售 4.7｜库存 95.3｜单价 50 → 净动销 **4.7**、理论 95.3、实际 95.3、盈亏 0、结算金额 **235.00 元**（修前为 250.00，差额归零）；5 组非半整数小数组合全部 exit 0，数量勾稽无误报 |
| N5 归档版本不对 | ✅ | **实测**：`archive/SKILL-被替换版-main.md` SHA-256 `4323ed14d726b279` == **main(c486788，894 行)**；「交付后必须回写台账（B5）」与 `writeback_ledger.py` 命令块 6 处命中；旧的 `SKILL-旧版-20261119.md` 已删除，archive/ 下只剩这一份 |
| N6 回归测试查错流致套件红 | ✅ | **实测**：断言已改为 `self.assertIn("净动销 4.7 件", run.stderr)`；全量 `python -m unittest discover -s tests` → **Ran 106 tests, OK**；`compileall` exit 0；测试残留目录 0 |
| N7 两引擎流语义不一致 | ✅ | **实测**：`monthly_recon_engine.py:35-41` 补上与铺货引擎相同的 `_REAL_OUT` / `emit`，SAVED／RESULT／WRITEBACK 三个交付行走 `emit`（stdout），过程输出走 stderr；铺货引擎 `:3044-3046` 同一通道，两个引擎现已一致 |

### 12.2 N6 的来龙去脉（保留记录）

`recon_engine.py:110-111` 新增了输出流重定向：

```python
_REAL_OUT = sys.stdout
sys.stdout = sys.stderr          # 过程中的一切输出走 stderr；失败时 stdout 保持 0 字节
```

过程输出改走 stderr 后，新加的 N4 回归测试仍断言在 stdout 上，因此一度红。**实测**该契约本身是好的：成功时 stdout 5 行（交付信息），失败时 stdout **0 字节**，便于机器侧判定成败；且无任何脚本依赖铺货引擎的 stdout（下游一律读 `_结算结果.json`／`_回写.json`／工作簿）。断言改到 `run.stderr` 后即绿。

### 12.3 闭环后的整体状态

| 段 | 条数 | 已修好 | 未落地 |
|---|---|---|---|
| P0 | 2 | 2 | 0 |
| P1 | 9 | 9 | 0 |
| P2 | 13 | 13 | 0 |
| P3 | 10 | 10 | 0 |
| 复核新增 N1–N7 | 7 | 7 | 0 |
| **合计** | **41** | **41** | **0** |

**验收**：`python -m unittest discover -s tests` → 106 项 OK；`compileall` exit 0；测试残留目录 0；仓库唯一新增文件为审计报告本身。

**剩余（非本 PR 阻断项）**：原报告第二节「结构性改进」六条——列定位、条码归一化、金额舍入、SHA-256 实现、复核包体积、主档解析——仍建议按原计划收敛。这六条正是 N3／N4／N7 这类「两套实现各自演化」问题的复发根源，建议另开任务处理，不必挡这次合并。
