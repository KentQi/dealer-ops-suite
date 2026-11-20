# PR #8 复核报告（独立复核 + 修复轮）

**复核对象**：`feat/structural-convergence`（PR #8，堆叠于 PR #1 之上，落地 issue #2–#7 六条结构性收敛）
**复核方式**：复核会话独立于实现会话，对交付记录逐条按证据实测；实现会话对复核发现逐条修复并复测。
**复核日期**：2026-11-20
**结论**：**实现成立，可以合**。复核发现的 F1（门禁覆盖面窄于自述）已在本轮修复；F2/F3/F4 一并收入本 PR。

---

## 一、复核方逐条核实（12 项，全部为真）

| | 复核方式与结果 |
|---|---|
| 基线 | `unittest discover` → 149 OK（106 基线 + 43），残留 0 |
| PR #8 | OPEN / base=`feat/dealer-ops-workflow` / MERGEABLE / mergeState CLEAN / 7 commits / 19 files |
| issues #2–#7 | 全部存在且 OPEN，标题与交付一致 |
| #2 列定位 | 复核方自跑新旧对拍：12/12 场景逐字节等价（返回值 + 错误文案） |
| #3 条码 | 16/16 形态全等（含 `.0` 串、float、bool、None、前后空白） |
| #4 舍入 | 7 个边界值行为改变（2.675→2.68、1.005→1.01、100.005→100.01…）；SKILL.md 迁移注记中 2.675→2.68 示例实测为真 |
| #4 rmb_upper | 22 个值中 20 个常规值逐字不变，变的 2 个（2.675、0.345）正是边界值本身——非回归 |
| #6 外置 | 侧车 + 纯文件名防路径穿越 + 字节 SHA 绑定 + 旧内联包兼容 + 无凭证不产侧车，5 条测试覆盖 |
| #7 主档 | `load_workbook` 主档 2 处 → 1 处 |
| golden vector | 23 项，覆盖别名命中/歧义阻断/泛化词守卫/等价类/预占冲突 |
| 文档 | SKILL.md 迁移注记、复核协议三态字段、故障与回滚新增两行，全同步 |

**复核方加分项记录**：`口径确认协议.md` 主动把 `koujing_sha256` 的过度声明改为事实（「口径哈希防护止于载入环节」）。

## 二、复核发现（F1–F4）

### F1【P1 · 已修复】去重门禁覆盖面窄于自述

门禁 docstring 写「任何脚本不得再私养一份」，实测两类绕过均绿：
1. `scripts/` 新放一个含自有 `file_sha256`/`norm_barcode`/`locate`/`money` 的新文件 → 门禁全绿（不扫全目录）；
2. 已覆盖文件注入 `hashlib.new('sha256')`（不含 `hashlib.sha256` 字面量）→ 门禁全绿（文本子串可规避）。

**修复**：门禁重写为 **AST 能力扫描**（`tests/test_no_duplicate_impls.py`）：
- 扫 `scripts/` **全目录**，新脚本自动纳管；
- 按**能力**判定：hashlib 别名、`hashlib.new("sha256")`、`hashlib.file_digest` 均命中；
- 违例签名 =「SHA 计算 + 流式喂食惯用法（iter-lambda / 循环 digest.update）或 file_digest」；纯内存原语摘要（findings_digest / values_digest / 附件摘要 / docx 逐部件指纹）不误伤；
- 自测 6 项：复核方的两类绕过场景原样种进临时目录必须被抓，两类合法语义必须不误伤；
- docstring 改为诚实边界声明（exec/getattr 级动态规避不在 AST 防御范围，靠 code review 兜底）。

**修复过程的额外收获**：精化后的门禁逐层暴露出原审计与复核都没列全的本地副本，全部收敛：
| 站点 | 性质 | 处置 |
|---|---|---|
| `plan_trust.file_sha256`（复核 F2） | 流式文件哈希第 5 份 | 收敛：`from schema_common import sha256_file as file_sha256`（`build_plan` 既有导入名不变）；`values_digest` 是取值指纹，保留 |
| `apply_fill.content_hash` | 流式文件哈希第 6 份 | 收敛（OSError→None 语义保留） |
| `make_todos.digest_file` | 流式文件哈希第 7 份 | 收敛（公开名保留） |
| `recon_engine` 条码 `.0` 剥尾 ×5 处（库存表 / 次月销售表 / 销售流水 / 批次表 / 促销使用表） | **第 4–7 套条码归一化**（原审计只列了三套） | 全部收敛到 `normalize_barcode`（守卫导入置于 `_bail` 定义前，覆盖所有调用点） |
| `verify_fill.package_fingerprint` | docx 逐部件内容指纹（合法语义） | 不收敛；据此精化门禁签名（「循环内 read」不算违例，「循环内 digest.update」才算），并在门禁内写明理由 |
| `apply_contract` 附件摘要 / `build_plan` values_digest / `locate_confirm` raw 摘要 | 内存原语调用 | 不收敛；门禁自测显式覆盖「不误伤」断言 |

最终：真实扫描零疑似违例；`.0` 剥尾惯用法在 `schema_common` 之外 grep 计数 0。

### F2【P2 · 已修复】plan_trust 第 5 份文件哈希实现

同上表首行。复核方确认 issue #5 原始清单不含 plan_trust、交付范围准确；本轮按「仓库层面单一实现」补齐。

### F3【P3 · 已修复】两处行为差异零文档

- **多列阻断**：SKILL.md §4.1 统一结算口径补「表头同一字段命中多列即阻断，不再任取左列——旧版行为已废除，原本能跑通的主档须删掉多余列后重跑（2026-11-20 行为差异）」；
- **`.0` 条码不再误报 P3**：README 政策表自检行（P 码唯一文档处）补「条码列带 `.0` 尾串的数字串按归一化处理，不再误报 P3（2026-11-20 行为差异）」。

### F4【P3 · 部分成立，已补钉】−34.7% 无回归防护

澄清：**「把 rule_result 塞回包内测试不会红」不准确**——`tests/test_recon_end_to_end.py:250` 对引用对象做**键集合全等断言**（`set(rule_ref) == {"_external","sha256","rows","business_type"}`），整包回填必然变红；月结 e2e 同理。真正没有防护的是**部分回填/旁路膨胀**。本轮补：
```python
packet_text = json.dumps(packet)
self.assertNotIn("quantity_reconciliation", packet_text)   # rule_result 独有结构
self.assertNotIn("priced_lines", packet_text)              # rule_result 逐笔行键
```
−34.7% 这个比例数字本身不钉进测试（比例随夹具形状漂移，脆断言），作为一次性基准记录于本节：1000 笔流水（含 10% 退货）夹具，包体 1,033,555 B → 675,207 B（−34.7%），侧车 332,467 B，确定性连跑逐字节一致。基准脚本在 `/tmp`（`v2c_bench/`），未入库。

## 三、修复轮后验证

- `python -m unittest discover -s tests` → **148 tests OK**（149 → 148：门禁两条旧文本断言合并为一条更强的接线断言，覆盖面不减反增）
- `compileall` 全部脚本通过
- 门禁自测 6/6：复核方两类绕过场景全被拦，两类合法语义全不误伤
- AST 真实扫描 `scripts/` 全目录：零疑似违例

## 四、第三轮：门禁召回加固（复核 D/C/E，同日）

复核方用自己的绕过代码（非门禁自测片段）重测修复轮门禁：F1 两类原绕过**全拦**（实质修复确认）；新探针发现三处能力盲区，本轮全部收口。

### D【P1 · 已修】金额舍入签名补 `round(Decimal(...), n)` 形态

`round(Decimal(str(v)), 2)` 用 HALF_EVEN，恰好复现 #4 要消除的口径分叉（复核实测：2.665 → 2.66 vs money 2.67；0.345 → 0.34 vs 0.35）。门禁 money 能力新增该形态（覆盖 `from decimal import Decimal` 与 `decimal.Decimal` 两种写法），植入自测钉死。
如实声明的边界：float 形态 `round(x, 2)` 与合法数量舍入（`_qval` 的 round(...,4)、T 数量键 round(v,2)）在 AST 层不可区分，不拦，靠 code review。

### C【P2 · 已修】列定位签名去掉参数名依赖

复核的 `locate(hdrs, keys)` / `pick_col(hdrs, wanted)` 语义等价重写绕过旧签名（依赖 alias/header 参数名）。新签名与参数名无关：**「enumerate/range(len) 循环 + 以检索词（本层/外层循环目标或函数参数）做 in 匹配 + 循环内 Return 索引」**。两个方向性约束防误伤（自测钉死）：
1. in 的**左侧**必须是检索词——`'合计' in line` 聚合行豁免扫描的检索词是常量，不命中；
2. 循环内必须有 Return——普通过滤/统计循环不命中。

**修复过程反查出第 8 处真残留**：`recon_engine._cc`（主档 SKU 段 keyword→列索引查找器，查条码/标准品名/别名/结算属性/规格/单位六列）——已收敛到 `locate_column`（精确优先、多列命中阻断；旧版子串先到先得废除，多列时 SKU 段中文中止）。其余 6 处旧命中（apply_v17／inspect_schema／make_master／make_todos／monthly `table`／scan_blanks）逐一验尸为合法域内逻辑（选工作表、渲染表格、聚合扫描），方向性签名下全部正确豁免。

### E【P3 · 已修（措辞）】一次性整读与表格宽度不一致

`sha256(open(p,'rb').read())` 是已声明不拦的取舍，但 docstring 表格写「文件 SHA-256 → 唯一权威」比实际签名宽。已改为「**流式**文件哈希」，边界段落显式把一次性整读列入已声明不拦清单，并新增自测把该取舍钉死（防止未来悄悄扩大签名时语义失配）。

### 验证

- 门禁自测 11 项全过（新增 5：Decimal-round 植入、enumerate 植入、range(len) 植入、一次性整读合法、聚合扫描合法）
- AST 真实扫描 `scripts/` 全目录：零疑似违例（含 `_cc` 收敛后）
- 全量 `python -m unittest discover -s tests` → **153 tests OK**
- 召回边界（如实声明，code review 兜底）：完全改名且换形态的等价重写、float 形态舍入、exec/getattr 级规避

## 五、遗留

- 门禁对 exec/getattr/动态字符串级规避无效（AST 的天然边界），依赖 code review——已在门禁 docstring 与本报告中明示，不另行「修复」。
- 门禁召回边界第 1、2 条（见第四节 D/C）同属签名式扫描的固有代价。
- `koujing_sha256` 若未来要落到复核包/复核锁，属新功能，不在收敛范围。
