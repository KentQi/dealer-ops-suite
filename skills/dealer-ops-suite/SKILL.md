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
8. **主档 A3 确认门槛不可跳过。** 普通候选只能预览；主档写入只接受确认页提交且带确认人的正式确认单，`--confirmed-by` 与 `--force` 均不能替代人工逐项确认。
9. **合同双门禁不可绕。** 写前必须 `--approve-write` 审阅；最终主合同必须有三文件 `released` 凭证并通过 `verify_fill.py`。
10. **待办只提醒，不执行责任认定。** 不自动算返利、扣保证金、认定窜货或创建定时任务；用户另行明确要求调度时才进入调度流程。
11. **对账后必须回写。** 先 dry-run，再 `--write`；未回写期间台账视为流程未完成。

详细结算口径以本文件 §4「统一结算口径」为唯一权威。

## 1. 意图路由

| 用户意图 | 进入 | 必需输入 |
|---|---|---|
| 审商务政策、查合同风险、合同包一致性 | 合同包审阅 | 主合同、防窜货附件、价格／出货政策 |
| 填合同、改已确认字段、出正式合同 | 合同编辑与双门禁 | 原始 docx 模板、累积全量 values、审阅记录 |
| 识别合同、建／更新客户台账 | A1–A5 主档人工确认链 | 合同原件；更新时还需已有主档 |
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
5. 用 `build_plan.py --spec … --doc <原始模板> --values <values.json> --out <全量-plan.json>` 把确认值编译成计划（`plan.json` 只能由脚本生成、失败时只留 `.rejected` 不得执行）；把“入参 → 落笔”对照表交用户确认。
6. 生成写前凭证，再填写；成品形成后生成三文件最终放行凭证并验收。

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

### 3.2 合同建档：A1–A5 人工确认链

LLM 负责抽取候选、保留原文依据、生成确认页和展示差异；用户在确认页逐项裁决。模型不得代填确认人、代点提交或把抽取结果当成用户结论。

```bash
# 仅没有主档时执行；已有主档必须跳过本命令，绝不重建覆盖
python3 "$SKILL_DIR/scripts/make_master.py" "<客户目录/_主档.xlsx>"

# A1 抽取候选：每字段带 value / evidence / confidence
python3 "$SKILL_DIR/scripts/extract_contract.py" "<合同文件>" -o "<候选字段.json>"

# A2 现状摘要：已有主档时用于核对；新客户可在 A5 再生成
python3 "$SKILL_DIR/scripts/make_digest.py" "<_主档.xlsx>"

# A3 生成并打开确认页；用户逐项确认、修改并填写确认人后提交
python3 "$SKILL_DIR/scripts/make_confirm_html.py" "<候选字段.json>" "<_主档.xlsx>" \
  -o "<合同确认页.html>"

# A3.5 用户回对话说「好了」后，自动定位正式确认单
python3 "$SKILL_DIR/scripts/locate_confirm.py" --customer "<客户目录>" --json
# 退出 0 时取 chosen.path；退出 2 有多份时不得静默任选

# A4 只将正式确认单写库：先 dry-run，再 --write；写前自动备份并归档
python3 "$SKILL_DIR/scripts/apply_contract.py" "<_主档.xlsx>" "<确认单.json>"
python3 "$SKILL_DIR/scripts/apply_contract.py" "<_主档.xlsx>" "<确认单.json>" --write

# A5 重跑摘要，供用户复核与归档；对账引擎直接读主档，不读摘要
python3 "$SKILL_DIR/scripts/make_digest.py" "<_主档.xlsx>"
```

**A3 确认门槛不可跳过。** 抽取结果是候选，不是结论；普通候选 JSON 只能 dry-run，不能 `--write` 落库，`--confirmed-by` 与 `--force` 均不能替代确认页。合同字段按页面最终确认的合同编号绑定；采用合同字段但未确认编号时，页面阻断提交。冲突由用户裁决，只有已确认单内的冲突覆盖才可另加 `--force`。

**A3.5 红线。** 不得要求用户移动、复制或重命名确认单。找不到时先确认是否已提交、是否换浏览器或电脑，再用 `--dir "<下载目录>"` 定位；多份候选时使用最新一份并明确告知，或请用户选择。

字段细节见 [references/客户主档设计.md](references/客户主档设计.md)。

### 3.3 待办

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

**统一结算口径（本文件为唯一权威，其他文档只引用不另述）：**

- 合同价格政策是货款单价唯一权威源；BIP 挂账导出只提供数量，不参与取价。业务类型必须显式登记或经用户确认，不根据文件名猜测。
- 价格与贴花按生效起止日管理；同一条码同一日期只有一条有效记录。断档、重叠、流水缺日期或数值非法时停止金额计算。合同附件未写生效期时，须经用户确认后才可按主合同期限生效；月中调价、新版政策必须有明确生效日。
- **金额逐笔先四舍五入到分，合计只汇总已舍入明细。** 销售、退货、签收、动销贴花与退货贴花同一规则，保证合计等于明细之和；不得先汇总未舍入金额再统一舍入。
- 退货／冲销数量在本期冲减，不修改历史对账单；金额凭原销售流水或明确登记的原价冲减，缺依据列为待补，不按退货日价格猜算。
- 盈亏列示金额＝盈亏数量 × 对账期末日生效合同价；正短缺、负盘盈，默认不计入结算总额。纳入须绑定本期起止日的单次授权，不继承、不取绝对值；不做货值恒等式，不倒挤调价价差。
- 贴花独立列示：正常销售按对账期末日标准；退货原销售已计提的按原销售标准本期同步冲减，未计提不冲，无法确认阻断为待补。盈亏不参与贴花。
- 月结：`买断货款＝Σ（签收数量 × 签收日合同价）－Σ（退货数量 × 原签收价）`，`月结贴花＝Σ（动销数量 × 动销日标准）－退货原贴花`。乙方销售流水只用于贴花；付款截止日不自行因节假日提前或顺延。

**最小验收清单（出账后逐项自查）：** 月中调价按两段价格分段计；短缺按期末价列示且默认不进总额；授权仅本期生效；任何非零盈亏均出单列示；价格断档／重叠／缺日期／退货缺原价均阻断；退货本期冲数量、按原售价与原贴花标准冲金额；月结签收买断不受乙方销量影响；直接买断进入对账即阻断；结果中不得出现比例阈值、绝对值盈亏或倒挤平衡项。

### 4.2 生成同一份配置并落盘

```bash
python3 "$SKILL_DIR/scripts/load_master.py" "<_主档.xlsx>" \
  --period 202610 --asof 2026-11-02 --cfg-out "<本期目录/RECON_CFG.json>"
export RECON_CFG="$(cat "<本期目录/RECON_CFG.json>")"
```

出账前已登记对账单编号时，在上述 `load_master.py` 命令追加 `--statement-no "<对账单编号>"`；未提供则保留空位，不自动编造或沿用其他期间编号。

- 指定期间目录不存在即停止，不回退最近月份或客户根目录。
- 铺货正式出账先完成 `_口径确认.json`；`--allow-legacy` 只供历史只读复核。
- 退货跨期追溯可传 `--original-sales <历史销售.json>`；月结还可传 `--original-receipts <历史签收.json>`。
- 需要签章 PDF 时在上述命令加 `--pdf`（可选 `--pdf-sheets "00,01,02,08"`，默认即这四张）；配置随 `--cfg-out` 原子落盘，供 v17 与 `export_pdf.py` 使用。
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

`apply_v17.py` 再次保存工作簿后，会同步 `_结算结果.json` 的路径与 SHA-256，保留原 `rule_result`。配置含 `pdf` 时按 `pdf_sheets`（缺省 00／01／02／08）调用 `export_pdf.py` 导出矢量签章 PDF（`<期间>对账-签章件.pdf`，宽度填满、高度切页、续页重复表头）；无 reportlab 时跳过导出并注明，不静默当作已出。月结不走 v17。

源表结构漂移先读 [references/数据映射与结构探查.md](references/数据映射与结构探查.md)，不改源导出、不猜列。

## 5. 复核、交付与回写

```bash
# 生成带业务规则底稿的复核包
python3 "$SKILL_DIR/scripts/make_review_packet.py" "<对账工作簿.xlsx>" \
  --cfg "<本期目录/RECON_CFG.json>" --master "<_主档.xlsx>" \
  --out "<复核包.json>" --md "<复核包.md>"

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
