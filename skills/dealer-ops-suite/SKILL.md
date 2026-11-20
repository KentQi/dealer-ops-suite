---
name: dealer-ops-suite
description: 经销商全生命周期工作套件。用于商务政策与合同包审阅、主合同／防窜货附件／价格政策放行、两层客户台账与履约待办，以及铺货代销和月结买断对账、贴花、复核与回写。触发语包括“审合同／识别风险／填合同／建客户档／合同台账／对账待办”“某月对账／算应结算／核对库存／出对账单／算贴花”“月结买断／签收买断／铺货代销／实销实结”“复核出账／验收／回写台账”，以及源表结构漂移、附表和签章 PDF 需求。直接买断不进入对账引擎；订单发货与签收在业务系统完成，不在本 Skill 内执行。
license: MIT
compatibility: Python 3.9+，openpyxl、xlrd、pdfplumber；可选 reportlab。合同无损写入需当前 harness 的 editor_sdk／doc_* 能力；L6 复核由当前模型执行，脚本不调用外部模型。
metadata:
  author: KentQi
  version: "2.0.0"
---

# 经销商全生命周期套件

本 Skill 服务真实业务顺序：

```text
商务条件谈判（在 Skill 外完成）
  → 确定出货政策、结算政策、指标与合同约束
  → 主合同＋防窜货附加协议＋价格／出货政策审阅、填写、放行
  → 正式合作后的发货、订单、签收（在业务系统完成，不在本 Skill 内）
  → 两层合同台账与本地待办
  → 铺货代销或月结买断对账
  → R1–R10 复核 → 交付 → 回写期间台账
```

`SKILL_DIR` 指本文件所在目录的绝对路径。命令统一写 `python3`；Windows 可用 `py -3`。路径使用绝对路径，不假定当前目录就是 Skill 或客户目录。

## 0. 不可突破的规则

1. **合同价格政策是货款单价唯一权威源。** BIP 只提供铺货代销的本期期初数量，不提供结算价格，也不要求与上期期末强制衔接。
2. **先识别业务模式。** 只让 `铺货代销·实销实结` 与 `月结买断` 进入各自引擎；`直接买断`立即阻断，不套公式。
3. **不猜缺失事实。** 价格断档／重叠、流水缺日期、退货缺原流水或明确原价、月结缺买断日／付款截止日，均停止金额计算。
4. **任何盈亏都出对账单。** 不设比例阈值；正数短缺、负数盘盈。盈亏按期末合同价列示，默认不进结算总额。
5. **盈亏纳入仅逐期单次授权。** 授权必须精确绑定本期起止日，不自动继承，不取绝对值。
6. **数量勾稽、金额分列。** 不恢复旧货值恒等式，不用调价价差或其他项目倒挤平衡。
7. **不替用户决定实质条款。** 不凭空填写目标、返利、保证金、违约金、账期或防窜货责任；LLM 只提示风险，不直接改合同。
8. **合同双门禁不可绕。** 写前必须 `--approve-write` 审阅；最终主合同必须有三文件 `released` 凭证并通过 `verify_fill.py`。
9. **待办只提醒，不执行责任认定。** 不自动算返利、扣保证金、认定窜货或创建定时任务；用户另行明确要求调度时才进入调度流程。
10. **对账后必须回写。** 先 dry-run，再 `--write`；未回写期间台账视为流程未完成。

详细结算口径见 [references/结算规则.md](references/结算规则.md)。

## 1. 意图路由

| 用户意图 | 进入 | 必需输入 |
|---|---|---|
| 审商务政策、查合同风险、合同包一致性 | 合同包审阅 | 主合同、防窜货附件、价格／出货政策 |
| 填合同、改已确认字段、出正式合同 | 合同编辑与双门禁 | 原始 docx 模板、累积全量 values、审阅记录 |
| 识别合同、建／更新客户台账 | 两层主档 | 已确认合同关键值及原文依据 |
| 生成对账、签回、付款、指标等待办 | 本地待办 | 主档；涉及工作日时需法定节假日及调休日历 |
| 铺货代销、实销实结、核库存盈亏 | `recon_engine.py` | BIP 期初、乙方销售／退货、乙方库存 |
| 月结买断、次月买断、按签收货款 | `monthly_recon_engine.py` | 签收、退货、乙方动销；买断日、付款截止日 |
| 直接买断 | 不进入对账 | 说明边界，不生成虚假对账结果 |
| 复核、验收、审计口径检查 | L6 R1–R10 | 当前工作簿、同目录 `_结算结果.json`、主档与配置 |
| 表名／表头／列顺序漂移 | 阶段 0 结构映射 | 原文件；用户逐项确认映射 |
| 试用装、中小样、效期、附表、签章 PDF | 铺货附表模块 | 已确认口径与现有铺货工作簿 |

已有客户、期间和文件时直接执行，不先问“想做什么”。能从合同或台账读到的事实不重复问；真正缺少的业务口径一次列清。

## 2. 商务政策与合同包

合同包固定由三份不同文件组成：

1. 主合同；
2. 防窜货附加协议；
3. 价格体系与出货政策（后续台账与结算的单价依据）。

### 2.1 审阅顺序

1. 用 `scan_blanks.py` 识别模板与待填字段；模板漂移即停止。
2. 收集用户确认值，形成累积全量 `values.json`，不自动补实质条款。
3. 运行确定性检查，再由当前模型阅读三份原文，产出带“文件＋条款＋原文”的 LLM findings。
4. 合并 findings，`blocker=0`，每个 `high` 按 finding id 登记 `status/note/owner`。
5. 生成写前凭证，再填写；成品形成后生成三文件最终放行凭证并验收。

写前审阅：

```bash
python3 "$SKILL_DIR/scripts/review_contract_package.py" \
  --spec "$SKILL_DIR/scripts/spec/jx.json" --values "<主合同-values.json>" \
  --fc-spec "$SKILL_DIR/scripts/spec/fc.json" --fc-values "<附件-values.json>" \
  --policy "<最终价格政策.xlsx>" --llm-findings "<llm-findings.json>" \
  --resolutions "<高风险处理.json>" --approve-write \
  --out "<合同写前审阅.json>" --md "<合同写前审阅.md>"
```

最终放行与主合同验收：

```bash
python3 "$SKILL_DIR/scripts/review_contract_package.py" \
  --spec "$SKILL_DIR/scripts/spec/jx.json" --values "<主合同-values.json>" \
  --fc-spec "$SKILL_DIR/scripts/spec/fc.json" --fc-values "<附件-values.json>" \
  --policy "<最终价格政策.xlsx>" --llm-findings "<llm-findings.json>" \
  --resolutions "<高风险处理.json>" --release \
  --main-file "<最终主合同.docx>" --fc-file "<最终防窜货附件.docx>" \
  --out "<合同最终审阅.json>"

python3 "$SKILL_DIR/scripts/verify_fill.py" \
  --template "<原始主合同模板.docx>" --result "<最终主合同.docx>" \
  --plan "<全量-plan.json>" --spec "$SKILL_DIR/scripts/spec/jx.json" \
  --values "<主合同-values.json>" --review-release "<合同最终审阅.json>"
```

写前凭证不要求尚不存在的成品哈希，但会绑定主合同／附件的 spec、values 与价格政策五类输入的绝对路径和 SHA-256；`apply_fill.py` 还核对本次 spec、values 正是审阅对象。最终凭证同时重新计算三份成品文件 SHA-256，并要求 `--result` 正是放行的主合同。审阅后替换任一输入或成品文件，验收均失败。

完整操作、无损编辑、回滚和纠错见：

- [references/合同包审阅.md](references/合同包审阅.md)
- [references/合同编辑.md](references/合同编辑.md)
- [references/field-map.md](references/field-map.md)
- [references/validation-rules.md](references/validation-rules.md)

## 3. 两层台账与本地待办

### 3.1 主档结构

`_主档.xlsx` 固定七表：

- **对账必需层**：主体、合同、价格、SKU、期间、贴花；
- **履约提醒层**：履约条款。

只登记关键值、原文依据、来源文件与位置，不搬运整份合同。合同按对账期间选择“完整覆盖本期的唯一合同”，零份或多份均阻断。价格与贴花按“条码＋生效起”保留历史版本，旧版本不得被新政策覆盖。

```bash
# 新主档
python3 "$SKILL_DIR/scripts/make_master.py" "<客户目录/_主档.xlsx>"

# 合同关键值写入：默认 dry-run，确认差异后才 --write
python3 "$SKILL_DIR/scripts/apply_contract.py" "<_主档.xlsx>" "<候选字段.json>"
python3 "$SKILL_DIR/scripts/apply_contract.py" "<_主档.xlsx>" "<候选字段.json>" --write
```

字段细节见 [references/客户主档设计.md](references/客户主档设计.md)。

### 3.2 待办

`make_todos.py` 从合同与履约提醒层生成 `_待办.json`，只写本地文件，不自动创建定时任务。

```bash
python3 "$SKILL_DIR/scripts/make_todos.py" "<_主档.xlsx>" \
  --period 202610 --calendar "<法定节假日及调休日历.json>" --out "<本期目录/_待办.json>"
```

日历格式：`{"years":[2026],"holidays":["2026-12-03"],"workdays":["2026-12-05"]}`。`workdays` 是调休上班日。无覆盖年份日历时，工作日待办标记 `estimated=true`，不给伪精确日期。付款始终保留合同原截止日，只有合同明确登记顺延／提前规则才调整。

## 4. 两种对账模式

### 4.1 输入与规则

| 模式 | 本期源文件 | 货款数量 | 正常取价 | 贴花 |
|---|---|---|---|---|
| 铺货代销 | BIP 期初、销售／退货、库存与必要回加 | 本期净动销 | 每笔销售日合同价；退货原销售价 | 正常销售期末标准；退货按原销售标准 |
| 月结买断 | 签收、退货、乙方动销 | 签收－退货 | 每笔签收日合同价；退货原签收价 | 乙方每笔动销日标准；不减少货款 |

铺货数量：`理论期末＝BIP 期初－净动销`，`实际结存＝库存导出＋回加`，`盈亏＝理论期末－实际结存`。

月结的乙方动销只计算贴花。买断日与付款截止日分别记录；直接买断不得进入任一引擎。

### 4.2 生成同一份配置并落盘

```bash
python3 "$SKILL_DIR/scripts/load_master.py" "<_主档.xlsx>" \
  --period 202610 --asof 2026-11-02 --cfg-out "<本期目录/RECON_CFG.json>"
export RECON_CFG="$(cat "<本期目录/RECON_CFG.json>")"
```

- 指定期间目录不存在即停止，不回退最近月份或客户根目录。
- 铺货正式出账先完成 `_口径确认.json`；`--allow-legacy` 只供历史只读复核。
- 退货跨期追溯可传 `--original-sales <历史销售.json>`；月结还可传 `--original-receipts <历史签收.json>`。
- 合同付款日无法识别时不默认次月 5 日；铺货可显式 `--pay YYYY-MM-DD`，月结必须在主档登记完整付款截止日。

### 4.3 执行引擎

```bash
# 铺货代销
python3 "$SKILL_DIR/scripts/recon_engine.py"

# 月结买断（二选一，不与铺货混跑）
python3 "$SKILL_DIR/scripts/monthly_recon_engine.py"
```

两者均输出工作簿、`_结算结果.json` 和 `_回写.json`。非结算类商品必须明确登记并被排除；未知属性阻断。缺原价、缺原流水、重复流水号或价格问题不会生成“估算总额”。

铺货需要 08／09 附表时：

```bash
python3 "$SKILL_DIR/scripts/apply_v17.py" --cfg "<本期目录/RECON_CFG.json>"
```

`apply_v17.py` 再次保存工作簿后，会同步 `_结算结果.json` 的路径与 SHA-256，保留原 `rule_result`。月结不走 v17。

源表结构漂移先读 [references/数据映射与结构探查.md](references/数据映射与结构探查.md)，不改源导出、不猜列。

## 5. 复核、交付与回写

```bash
# 生成带业务规则底稿的复核包
python3 "$SKILL_DIR/scripts/make_review_packet.py" "<对账工作簿.xlsx>" \
  --master "<_主档.xlsx>" --out "<复核包.json>" --md "<复核包.md>"

# 当前模型按 R1–R10 生成报告后
python3 "$SKILL_DIR/scripts/apply_review.py" "<对账工作簿.xlsx>" "<复核报告.json>" \
  --packet "<复核包.json>" --apply
python3 "$SKILL_DIR/scripts/apply_review.py" "<对账工作簿.xlsx>" --check

# 台账回写：先 dry-run，再明确写入
python3 "$SKILL_DIR/scripts/writeback_ledger.py" "<_主档.xlsx>" \
  --from-json "<本期目录/_回写.json>" --json
python3 "$SKILL_DIR/scripts/writeback_ledger.py" "<_主档.xlsx>" \
  --from-json "<本期目录/_回写.json>" --write --json
```

复核包优先读取同目录 `_结算结果.json`：工作簿路径或当前哈希不一致即拒绝，主档与各业务源表也必须匹配出账时路径及 SHA-256，不沿用旧结果。铺货复核数量勾稽、逐笔销售／退货金额、盈亏列示和单次授权；月结复核签收净额、原签收价退货及动销贴花。月结不把写回用的 operational zero 当作 BIP 期初。

`--apply` 必须提供复核包；任一 R1–R10 清单项不通过时，verdict 不得标为「通过」。正式回写只接受结果凭证内嵌的 `_回写.json`，并核对目标主档仍是出账快照、复核锁 `pre_sha256` 等于出账工作簿哈希、当前工作簿哈希等于 `post_sha256`，且归档报告与锁结论一致。完全相同的重复回写可幂等跳过。

完整协议见 [references/复核协议.md](references/复核协议.md)。任何复核后保存都会使复核锁失效，必须重评。

## 6. 边界与故障

- 发货、订单、乙方签收在业务系统执行；本 Skill 只消费导出结果。
- 盈亏不等于责任认定，返利／保证金／窜货不由待办自动执行。
- 净退货可产生负结算数量、负金额和负贴花，回写保留符号；不得为通过非负校验而改数。
- 无乙方库存的 SKU 标为待补，不按库存 0 计算盈亏。
- 合同审阅不是法律意见；对账复核不是审计意见。
- 报错、配置字段、备份与回滚见 [references/故障与回滚.md](references/故障与回滚.md)。
