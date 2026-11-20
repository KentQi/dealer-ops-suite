# 任务队列

<!--
行格式(解析器只认以下三种任务行,两空格+括号勿动):
  - [ ] T004 任务名                              ← 未认领
  - [~] T003 任务名  (model, YYYY-MM-DD HH:mm)       ← 已认领(relay_claim 标注)
  - [x] T002 任务名  (YYYY-MM-DD)                        ← 已完成(relay_session_end 盖戳;完成提交反查 git log --grep "T### 完成")
任务行下一行可写验收(完成判定,二选一):
  验收: <一条可执行命令>
  验收: test:<失败测试文件路径>   ← 红灯协议:交班者留下失败测试,接班者让它变绿
启用 worktree 时,relay_claim 返回 controlRoot(协议控制面)与 executionRoot(代码执行面);
executionRoot 是独立检出(node_modules 等依赖需单独安装)。任务分支完成不等于已合并:
出场绿闸双提交(任务分支代码 + controlRoot 的 TASKS/LOG),分支由人工 merge,过期缓存清理保留分支。
下方为示例条目,首个真实会话后替换为真实任务。
-->

## 进行中
(空——认领后由 relay_claim 移入;认领锁在 .relay/claims/,TTL 4 小时,崩溃会话的认领到期自动可接管)

## 待办
- [ ] T011 出账流水线强制禁词扫描（新运行时行为）
  验收: apply_v17 附表保存后自动跑 check_public_copy，违禁词退出1并报「对外文案门禁未通过」；e2e 负路径测试覆盖；DECISIONS 落条目
- [ ] T006 待用户指定后续方向
  验收: 无需强制验收；若用户要求可按 SPEC AC1–AC4 全量巡检
  备注: 全仓库实质遗留已清零（审计 41 项、结构性收敛 6 条、复核 F/D/C/E、文档漂移、koujing_sha256 全部闭环）；后续工作由用户指定

## 已完成
- [x] T010 最小CI建立（恢复批次4）  (2026-11-21)
  验收: .relay 白名单含 .github；tools/validate.py 迁入（.claude-plugin 降级可选）；.github/workflows/ci.yml 存在（pytest统一、SHA钉、单版本、gitleaks）；GitHub Actions 首绿
- [x] T009 迁移对抗测试层与六类回归钉子（恢复批次3）  (2026-11-21)
  验收: fill_sim/fixtures/test_attacks/test_workflow 迁入（真实放行凭证夹具、品牌scrub）；test_public_copy_gate/test_v17_gates/test_deterministic_rerun/铺货侧映射测试新增；pytest 全绿且存量 unittest 不回归
- [x] T008 锁定表静默no-op显式披露、文档收口与L6术语退役（恢复批次2）  (2026-11-21)
  验收: verify_fill 在 locked_content 为空时输出「未登记」披露文案；validation-rules/field-map/合同编辑 三处宣称收口；L6 术语退役（README 2 处 + 口径确认协议 2 处 + 复核协议 2 处 + 对外文案治理 2 处）；pytest 全绿
- [x] T007 恢复依赖声明与requirements（恢复批次1）  (2026-11-21)
  验收: SKILL.md frontmatter 与 README 依赖声明含 python-docx/lxml；根目录 requirements.txt 存在且 pip install -r 可解；pytest tests -q 全绿
- [x] T005 实现 koujing_sha256 复核链路绑定  (2026-11-21)
  验收: SPEC AC1–AC4 全绿（153+ tests）；篡改场景测试：出账后改 _口径确认.json → apply_review --check 拒绝并给中文报错；口径确认协议.md §6.3 与 koujing.py docstring 收回如实声明改为实现态
  备注: 裁决已落 DECISIONS（2026-11-21 现行）。实现面：两引擎写 koujing_sha256 进 _结算结果.json 凭证 → make_review_packet 携带 → apply_review --check 比对凭证与当前口径文件；k oujing.py docstring 与协议文档同步
- [x] T004 用户裁决 koujing_sha256 是否落地  (2026-11-21)
  验收: DECISIONS.md 新增 koujing_sha256 裁决条目（落地→另开实现任务按 SPEC AC1–AC4 判定；不做→状态现行并删除 SPEC §4 对应行）
  备注: SPEC §4 已挂该开放问题及背景（口径确认协议曾承诺、PR #8 时已改为如实声明）；此为用户裁决项，非代码任务
- [x] T003 按需领取后续小任务  (2026-11-21)
  验收: 候选任务清单见 docs/reviews/ 两份报告的「遗留」节；无强制验收，取任意一条落地并按 SPEC AC1–AC4 判定
  备注: 仓库现仅剩 main（1f40188），153 tests OK，Pages 与源图一致；relay 协议文件中 BIP 提及均为政策性语境（验收行/日志/决策），非内容回潮
- [x] T002 合并后 Pages 构建确认  (2026-11-21)
  验收: gh api repos/KentQi/dealer-ops-suite/pages/builds/latest -q .status 为 built 后，抽查 docs/architecture.html 与 docs/business-workflow.html 与 docs/source/*.json 一致（grep 关键节点标签），且内容级 BIP 扫描维持零命中
  备注: Pages 由 /docs 构建合并后自动触发；图源为 docs/source/*.json（9160aff 重绘版）
- [x] T001 人工合并 PR #9 并确认合并后基线  (2026-11-20)
  验收: 合并后本地跑 python3 -m unittest discover -s tests（预期 153 OK）；内容级 grep -rn BIP --include=*.py --include=*.md . 过滤对外文案词表后为 0
  备注: PR #9 分支 fix/public-copy-sync（head 08a900b）已 rebase 到 9160aff，图与 plugin.json 与 main 零差异，只含脱敏+守卫+文档；worktree .worktrees/fix-public-copy-sync 保留待合并后清理
(空——relay_session_end 盖完成章后由引擎移入)
