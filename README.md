# Dealer Ops Suite · 经销商运营全流程套件

把「谈客户 → 定政策 → 签合同 → 建台账 → 月度对账」串成一条可执行的工作流。

> **公开脱敏版**：仓库中的公司、经销商、银行、税号、合同编号、金额与商品均为虚构示例。使用前请替换为你自己的合同模板、字段清单和业务口径。

## 这个项目解决什么问题

经销商业务里最耗神的不是算一次账，而是同一套信息在合同、客户主档、台账、对账单之间反复搬运：

- 合同里的主体信息、付款日、异议期、签回期，每次对账都要重新翻；
- 客户换一个导出模板，表名、列顺序、字段名就全变；
- 对账单算完还要人工复核，口径散落在聊天记录和 Excel 备注里；
- 政策表里同一条码出现多行、行内门槛和表尾备注打架，没人先发现。

Dealer Ops Suite 把这些问题收进两条完整工作流：**填合同**与**月度对账**。它不替业务做决定，只负责把已确认的信息准确落笔、把该拦的错误拦住、把该留的痕迹留下。

## 能力

| 模块 | 做什么 | 主要入口 |
|---|---|---|
| 合同填充与条款校验 | 识别经销 / 铺货代销合同及附件待填项，一次性收值、关键条款校验、编译计划、写入、验收 | `scan_blanks.py`、`build_plan.py`、`apply_fill.py`、`verify_fill.py` |
| 客户主档 | 从合同抽取候选字段，生成确认页，人工逐项确认后写入客户主档 | `extract_contract.py`、`make_confirm_html.py`、`apply_contract.py` |
| 月度对账 | 主档转配置，计算应结算、差异、贴花，生成对账工作簿 | `load_master.py`、`recon_engine.py` |
| 复核与交付 | L1–L5 确定性校验 + L6 LLM 语义复核，复核锁绑定工作簿哈希 | `make_review_packet.py`、`apply_review.py`、`check_public_copy.py`、`export_pdf.py` |
| 政策表自检 | 只做三类原则性检查：同条码多行、行内与表尾门槛冲突、条码缺失或非法 | `check_policy.py` |
| 台账待办 | 按台账公式计算当前应做、距截止、紧急度；公式无缓存时即时重算并标注 | `ledger_read.py` |
| 台账回写 | 对账交付后把期初 / 结算 / 差异 / 贴花 / 状态写回主档「期间」表；dry-run、自动备份、同期间幂等 | `writeback_ledger.py` |

### 合同关键条款校验（填写阶段）

`build_plan.py` 与 `verify_fill.py` 会在写入前后检查：

- 合同编号格式、类型码、年份与签订时间一致性；
- 合同起止、签订时间、附件签署日期；
- 履约保证金大小写是否成对、是否一致；
- 季度/年度返利区间，以及年度返利是否明显高于季度；
- 季度/月度销售目标是否至少给出一档、分解合计是否提示复核；
- 主体名称、税号、联系人、地址、电话格式；
- 模板漂移、空白未填、锁定条款表被改、落笔值与确认值不一致。

这些是**确定性校验**，解决“填错/漏填/前后矛盾”。条款语义风险（单方义务、定义缺失、罚则不可执行、自动续约/解除冲突）不在脚本内，需由 LLM 审阅并逐条让用户确认；目前这部分还没有独立模块。

## 两张图

- [架构图](https://kentqi.github.io/dealer-ops-suite/architecture.html)：入口、插件清单、SKILL.md 路由、四类能力、支撑资源与交付物。
- [业务流程图](https://kentqi.github.io/dealer-ops-suite/business-workflow.html)：谈客户 → 定政策 → 合同与资料 → 签署 → 订单发货（Skill 外）→ 客户销售 → 建台账 → 生成待办 → 对账。
- [项目首页](https://kentqi.github.io/dealer-ops-suite/)：两张图的入口。

也可以下载 `docs/architecture.html` 与 `docs/business-workflow.html`，用浏览器离线打开。

## 目录

```text
.
├── plugin.json
├── README.md
├── LICENSE
├── docs/
│   ├── index.html
│   ├── architecture.html
│   ├── business-workflow.html
│   ├── SANITIZATION.md
│   └── source/
│       ├── architecture.json
│       └── business-workflow.json
└── skills/
    └── dealer-ops-suite/
        ├── SKILL.md
        ├── assets/
        │   ├── 客户主档-模板.xlsx
        │   └── copy/对外文案.json
        ├── references/
        │   ├── 客户主档设计.md
        │   ├── field-map.md
        │   ├── 复核协议.md
        │   ├── 口径确认协议.md
        │   ├── 故障与回滚.md
        │   ├── 对外文案治理.md
        │   ├── 数据映射与结构探查.md
        │   ├── confirm-template.md
        │   └── validation-rules.md
        └── scripts/
            ├── spec/
            │   ├── jx.json
            │   ├── dx.json
            │   └── fc.json
            └── 26 个 Python 脚本
```

## 快速开始

> Skill 指令本身不绑定某个 harness。下面以 WorkBuddy 的目录为例；换到其他 harness 时，只需把 `skills/dealer-ops-suite` 放进它的 skills 目录。Office 写入若没有无损编辑通道，Skill 会明确提示能力边界，不会用 python-docx 强改。

### 方式一：作为 Agent Plugin

```bash
git clone https://github.com/KentQi/dealer-ops-suite.git \
  "<你的 harness 插件目录>/dealer-ops-suite"
# WorkBuddy 示例：~/.workbuddy/plugins/dealer-ops-suite
```

重启会话后，在对话里说「用 dealer-ops-suite 帮我识别这份合同」或「用 dealer-ops-suite 做 2026-10 对账」即可。

### 方式二：直接作为 Skill

```bash
git clone https://github.com/KentQi/dealer-ops-suite.git /tmp/dealer-ops-suite
cp -R /tmp/dealer-ops-suite/skills/dealer-ops-suite \
  "<你的 harness skills 目录>/dealer-ops-suite"
# WorkBuddy 示例：~/.workbuddy/skills/dealer-ops-suite
```

### 依赖

```bash
python3 -m pip install openpyxl xlrd pdfplumber
# 可选：导出矢量 PDF
python3 -m pip install reportlab
```

要求 Python 3.9+。脚本不联网，不调用外部模型；L6 语义复核由当前 Agent 的模型执行。

## 配置成你自己的业务

公开版刻意把品牌相关的固定值全部换成了示例。真正使用前，至少改这几处：

1. **合同模板**：把 `示例品牌经销合同-JX.docx`、`-DX.docx`、附件一放到你自己的 `经销合同模板/` 目录。
2. **字段清单**：改 `skills/dealer-ops-suite/scripts/spec/jx.json`、`dx.json`、`fc.json`，让 `anchor`、`table`、`row`、`col` 与你的模板一致。
3. **甲方固定信息**：改 `spec/*.json` 的 `readonly_check`，换成你自己的主体名称、统一社会信用代码、开户行与账号。
4. **品牌码与编号规则**：公开版用 `DM` 作为示例品牌码。改 `build_plan.py` 与 `spec/*.json` 里的示例编号。
5. **客户主档模板**：改 `assets/客户主档-模板.xlsx` 的表头与示例行。
6. **对外文案规则**：改 `assets/copy/对外文案.json` 与 `references/对外文案治理.md`，按你的合规要求增删禁用词。
7. **附件一锁定内容**：公开版把 `fc.json` 的 `locked_content` 置空。请从你自己的附件一模板复制标准条款表，否则 `verify_fill.py` 不会检查这一块。

改完后建议先拿一份**空白模板**跑：

```bash
python3 skills/dealer-ops-suite/scripts/scan_blanks.py \
  --spec skills/dealer-ops-suite/scripts/spec/jx.json \
  --doc "/path/to/你的经销合同.docx"
```

## 安全与隐私

这是一个**脱敏公开版**，不包含任何真实客户、合同、价格政策或对账数据。

- 所有公司名、经销商名、地址、税号、银行账号均为虚构示例；
- 合同编号统一使用 `DM-` 示例品牌码；
- 附件一的具体违约金标准已移除，改为「按合同约定」；
- 未包含真实合同模板、客户主档、对账工作簿、政策表；
- `.gitignore` 默认忽略 `*.xlsx`、`*.docx`、`*.pdf`、`values.json`、`*_主档*`、`*对账*` 等常见业务文件。

如果你要在此基础上接入真实业务，**不要把真实数据提交到公开仓库**。建议：

1. 私有仓库保存模板与客户数据；
2. 公开仓库只保留代码、文档与虚构样例；
3. 提交前跑一次 `gitleaks detect --no-git`；
4. 确认 `values.json`、客户主档、对账单没有被 `git add`。

## 设计原则

- **能从文件读到的事实，不提问**；必须问的，一次批量问完。
- **缺失即阻断**：空值不会被默认成「结算类」「0 件」「件」。
- **不替用户做业务决定**：返利、保证金、销售目标、兑付方式，用户给多少填多少。
- **不覆盖模板**：永远另存为新文件，模板只读。
- **复核锁绑定工作簿哈希**：保存一次就失效，必须重新复核。
- **对账结果必须回写台账**：没有写入主档「期间」表，就不算本次对账完成。
- **对外单据不泄露内部口径**：内部系统名、费用目的、出数规则不得进入对账单。

## 边界

- 订单与发货不在本 Skill 范围内，业务图里只作为交接节点。
- 合同填充依赖你提供的 `.docx` 模板；模板改版后必须同步 `spec/*.json`。
- 对账依赖甲方挂账、库存、销售流水三份结构化导出；扫描件不做 OCR 兜底。
- 政策表自检只覆盖三类原则性问题，不做跨客户比价、毛利测算或折扣建模。
- L6 语义复核需要 Agent 具备文件读写与推理能力；脚本本身不调用模型。

## License

MIT。见 [LICENSE](LICENSE)。
