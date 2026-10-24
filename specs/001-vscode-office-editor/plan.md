# Implementation Plan: VS Code 原生聊天驱动的 Office 文档编辑

**Branch**: `main`（实际 Git 分支）；Spec Kit feature ID 为 `001-vscode-office-editor` | **Date**: 2026-10-21 | **Spec**: [spec.md](spec.md)

**Input**: `specs/001-vscode-office-editor/spec.md`，沿用 Excel 优先的范围假设。状态：Phase 0/1 设计完成，未开始实现。setup-plan 根据 feature.json 返回的 BRANCH 是逻辑 feature ID，本轮没有创建或切换 Git 分支。

## Summary

新增桌面 VS Code workspace extension：在自定义编辑器中复用表格渲染，允许手工及原生 Agent 修改同一工作簿；复杂分析由原生终端运行工作区 Python，以分型文件交换输入和结果。复用现有 xlsx gateway、Rust sidecar 与 Univer 组件，新增独立于 Electron 的权威会话和窄宿主端口。

关键新增工作是 host 权威状态、隔离候选事务、持久可序列化历史和不重建会话的保存；不能只包装现有 App/CLI，因为现有执行器可能部分应用，现有跨保存撤销也只覆盖部分情形。详见 [research.md](research.md)。

## Technical Context

**Language/Version**: TypeScript 5.9 系列、ES2022；开发 Node >=22.12/npm >=10，扩展宿主 Node 22 兼容；Rust edition 2024/Cargo.lock；Python 3.11+ 为验收环境，不作为插件内嵌运行时。

**Primary Dependencies**: VS Code 稳定 API 基线 1.109.5、React 19、Univer 0.25.1（共享同一解析版本）、Zod 4、`@genoffice/xlsx-gateway`、现有 xlsx-sidecar；新增构建使用 esbuild/Vite，测试使用 Vitest、Playwright、`@vscode/test-cli`/`@vscode/test-electron`。新增依赖选择遵循仓库许可检查。

**Storage**: 工作簿继续为 .xlsx；扩展私有目录保存打开快照、增量状态、历史、备份及事务收据；workspace `.genoffice/work/` 为用户可查看的程序输入、脚本及结果，加入工作区排除建议但不静默修改 Git 配置。无数据库、无网络服务。

**Testing**: 纯会话/契约 Vitest，Rust 现有测试；浏览器组件 Playwright；VS Code 集成宿主验证真实编辑器生命周期、工具及 VSIX；人工原生聊天验收单独记录。错误注入覆盖 prepare、commit、save、recovery 全边界。

**Target Platform**: VS Code >=1.109.5 桌面；首版验证 linux-x64/Ubuntu 22.04、win32-x64/Windows 11、darwin-arm64/macOS 14，Remote SSH 首先 Ubuntu 22.04 x64。workspace host 负责二进制、文件及 Python，client Webview 负责界面。纯浏览器、虚拟文件系统、不在矩阵的架构不声明支持。

**Project Type**: Monorepo 内新增桌面编辑器扩展与共享库，Electron 产品保留独立宿主适配。

**Performance Goals**: 规格 SC-001 的 100k-cell 两表首屏 19/20 次 <=10 秒；SC-005 的校验完成后 1k 普通单元格应用 19/20 次 <=2 秒；100k 输入 Python 汇总正确。设备、数据、冷暖启动和计时点见 quickstart，均为目标而非已测结果。

**Constraints**: 不实现新聊天/模型循环；稳定 API，不监听私有聊天事件；保存前后保留历史；每失败批次无部分提交；只在执行前保证可撤销时接受修改；未经支持的特性保持原包或阻止有损保存。样式按语义 token，文档原有颜色不受宿主主题污染。

**Scale/Scope**: .xlsx；常用单元格/公式/格式、工作表、基本图表；聊天读 2k 格；展开写入每批 2k 格，导入默认 1k 格批次；范围操作最多 200k 格且通过历史预算；文件交换每次 <=1m 格且 <=64 MiB；单事务双向变化 <=64 MiB，保留历史 <=256 MiB；OOXML 条目 500 MiB 限制沿用。

## Constitution Check

*GATE: Phase 0 之前及 Phase 1 之后检查。*

| 检查 | Phase 0 前 | Phase 1 后 |
| --- | --- | --- |
| 正式宪章原则 | 宪章为未填模板，无可执行原则，不据示例虚构约束 | 状态不变，本轮未代填宪章 |
| 当前授权边界 | 只做规划，不写业务代码、不发布 | 仅生成设计文档，未生成 tasks 或实现 |
| 用户价值与范围 | 原生聊天、技能/Python、手工+AI均入规格；Excel优先为显式假设 | 28 条 FR 均映射至设计与验收 |
| 主题/i18n/现有行为 | 遵循 CLAUDE.md，不修改文档数据色值 | 宿主主题只映射 UI token；新增 UI 字符串使用共享 i18n |
| 数据一致性 | 明确当前原子性与历史缺口 | 隔离候选+host commit、版本/收据、非重开保存已定义 |
| 宿主依赖隔离 | 不在 Webview/公共域层加载 Electron/Node | 窄端口、浏览器/Node 构建边界及回归门槛明确 |

结果：通过可适用的检查，无需合理化的宪章例外。正式宪章尚未制定是治理事实，不当作本次已批准治理条款。

## Project Structure

### Documentation (this feature)

```text
specs/001-vscode-office-editor/
├── specify-prompt.md
├── spec.md
├── plan.md
├── research.md
├── data-model.md
├── quickstart.md
├── checklists/requirements.md
└── contracts/
    ├── document-session.md
    ├── agent-tools.md
    └── python-exchange.md
```

`tasks.md` 留给 speckit-tasks，本轮不创建。

### Source Code (repository root)

以下为拟新增/抽取的结构，不代表文件已经存在：

```text
apps/vscode-office/
├── package.json                 # VSIX、工具、命令、custom editor 配置
├── src/extension/               # provider、tools、commands、resource/backup adapters
├── src/webview/                 # 无聊天的表格编辑入口、host bridge、主题适配
├── skills/genoffice-workbook/   # SKILL.md 与操作/文件交换说明
├── resources/sidecar/           # 构建时按 target 放置二进制
├── scripts/                    # 打包、夹具、性能与 Extension Host 测试入口
└── tests/                      # 集成、工具和实际 VSIX 验证
packages/workbook-session/
├── src/                        # state、transactions、history、save、exchange
└── tests/                      # 版本、隔离、恢复、并发、保真、预算
packages/sheets-editor/
├── src/                        # 浏览器组件、intent adapter、candidate runtime
└── tests/                      # 手工命令、投影重建、候选资源释放
packages/xlsx-gateway/           # 现有保真引擎，增补完整状态序列化边界
apps/sheets/src/                 # Electron 适配器；旧导入转导出以渐进迁移
apps/sheets/native/xlsx-engine/  # 现有 Rust 引擎继续复用
```

**Structure Decision**: 扩展放入既有 `apps/*` workspace；新增会话层是 Node 库，编辑组件是浏览器库，两者不互相导入宿主。`@genoffice/ui` 当前入口会连入 AI 依赖，因此只复用经验证的独立组件/CSS 子入口，避免把 agent-core/ai-provider 装进新 Webview。不复制整个 App.tsx，不用空实现伪造 DesktopApi。

## Phase 0 — Research Complete

研究及选择见 [research.md](research.md) R1–R10。没有未解决的架构澄清项。稳定 API 的限制以显式绑定和 workspace 技能安装解决；原子执行与历史的工程风险用前置原型及失败门槛验证，不声称已有实现。

## Phase 1 — Architecture and Contracts

### 权威模型、投影与手工操作

每 URI 一个 WorkbookDocument 和独立 sessionId。Node 会话拥有 originalBase、typed overlay、结构日志、timeline；编辑组件拥有当前可视投影。手工支持命令在执行前拦截为 intent；按键/输入法输入中的临时文本不当作已经提交内容，读取前先结束可提交编辑或返回 editor_busy。AI、导入和手工全部经同一队列。

候选操作在隔离 runtime 中基于固定 stateId 执行并产生可序列化 changeset；host 验证版本、约束、历史容量，持久提交后再发布视图。未注册的可变命令禁止执行，保护/公式依赖/图表等变更必须纳入 changeset。投影失效时从 host 重建，无可见 Webview 时 host 仍可读取与保存；需要候选解释器但暂不可用时返回明确错误，不回退旧文件。

首版允许的手工及 AI 能力：值/基础公式，font/bold/italic/颜色/fill/number format/对齐/wrap/border，区域填充与复制，sheet add/rename/delete（引用保护），row/column insert/delete（已有 gateway 校验），column/bar/line/pie 图表。未知特性保持只读及原包，不承诺完整 Excel 等价或任意宏执行。是否可执行以 context 返回 capabilities 为准。

### VS Code 生命周期

实现 open、resolve、save、saveAs、revert、backup、dispose；每个内容事务发一个 CustomDocumentEditEvent，其 undo/redo 回调移动 host timeline，并通知投影。Webview 内快捷键最终使用同一 VS Code 历史入口，内部候选历史不构成第二条用户历史。保存同队列排他执行，序列化 originalBase+完整当前状态，所有阶段先在私有目标完成，再保真校验及冲突检查、替换用户目标，成功后更新 savedStateId/disk hash。保存不重新打开当前 runtime，不剪裁 history。

source URI 与 sidecar 本地路径区分；saveAs 只写目标，不私自重写只读 URI，若 VS Code 随后建立新文档则创建新 epoch。revert 使所有 binding 失效并重新载入，历史按 VS Code revert 生命周期重置。backup 使用私有目录可完整恢复 base/overlay，异常恢复生成新 session，恢复点明确；不承诺恢复备份成功之前的编辑。

宿主原子替换依赖 POSIX rename 或 Windows 可恢复替换，保存失败保留原/备份及当前内存状态；哈希检查不能构成对任意第三方写进程的全局互斥，发布前后与 watcher 均检查，检测竞态保留 staged/original 副本并进入 conflict，不虚称跨进程 CAS。该限制在用户提示与测试中明确。

### 原生聊天与程序通道

注册稳定 LM tools，不实现新 Chat Participant。工具 prepareInvocation 只描述和确认，不修改；invoke 中重新校验目标、revision、文件和取消状态。绑定通过命令冻结选区，或 context 明确 URI；从工具返回 bindingId 后必须显式传递。原生聊天发送瞬间自动采集不是首版承诺。

Python 由原生 Agent 终端运行。export 输出当前 typed 快照；prepareImport 对输入文件制作私有不可变副本；applyImport 以 planId+cursor 分批。每个批次原子且可撤销；取消只取消未提交部分。用户在批间修改将产生 revision_conflict，旧计划终止并需要重新计算，不通过静默刷新 expectedRevision 继续套用旧结果。具体契约见 [agent-tools](contracts/agent-tools.md)、[python-exchange](contracts/python-exchange.md)。

### 构建、远程与技能

将 Node sidecar client 和保存序列化服务迁出 Electron main，旧路径转导出防止复制；不得把 Electron 的 process.resourcesPath 用在扩展。每个平台 VSIX 带匹配的 sidecar，无运行时下载和 cargo 要求；开发阶段构建 Rust。`extensionKind=workspace`，Remote SSH 的快照/Python/sidecar 均在远端，Webview 资源用 asWebviewUri、限定 localResourceRoots、nonce CSP，禁止任意执行消息；二进制不支持或 workspace 非可信时不启动会话及工具，给出能力说明。

提供显式 Install Workbook Skill 命令，目标为选定 workspace 的 `.github/skills/genoffice-workbook/`；已有内容先比较再由用户选择。首版不依赖尚未在最低版本实测的 chatSkills 打包路径，不修改用户模型配置。原生工具审批由 VS Code 承接。

## Implementation Sequence and Verification Gates

这是一组交付顺序，不是可执行 tasks 清单：

1. **G0 核心风险原型**：固定版本验证 CustomEditor/LM tools；手工命令截取与 offscreen candidate；“结构+图表—保存—撤销—重做”状态可序列化；故障注入不污染主状态。失败时修正隔离实现，不允许以非原子或丢历史代替规格。
2. **G1 共享库抽取**：窄端口、sidecar client、schema、保真序列化、独立 UI 入口；Electron 回归及无 Electron 的扩展构建都通过。
3. **G2 文档完整闭环**：host timeline、手工操作、保存/另存为/冲突/恢复，先达到 US1/US4/US5。
4. **G3 原生 AI 与技能**：稳定工具、目标绑定、read/apply、权限描述、workspace 技能安装，达到 US2/US6。
5. **G4 Python**：typed snapshot/result、prepare/apply、分批和取消，达到 US3，证明100k输入不经过聊天正文。
6. **G5 发布验证**：三平台 VSIX、最低/当前稳定 VS Code、本地与 SSH、确定性回归及真实聊天检查；达标前不宣传容量/性能。

### Requirement Coverage

| 需求 | 设计/契约 | 验证入口 |
| --- | --- | --- |
| FR-001–003 | CustomEditor、共享显示、save/backup | quickstart V1、V6 |
| FR-004–009 | binding、read、apply 工具 | V2、V5 |
| FR-010 | workspace 技能安装、更新差异 | V2、V6 |
| FR-011–016 | typed exchange、prepared import、当前状态导出 | V3、V4 |
| FR-017–020 | revision、epoch、串行队列、candidate commit | V4、V5 |
| FR-021–024 | host timeline、历史预算、backup、分批状态 | V1、V4、V5 |
| FR-025–028 | 保护、能力、保真校验、idempotency | V4、V5、V6 |

## Complexity Tracking

没有正式宪章例外。新增两个共享包分别对应不同运行环境及生命周期；权威状态/candidate/history 是满足规格的必要复杂性，不能简化为磁盘 CLI 覆盖。首版限制为 .xlsx、单文档单编辑视图及明确操作白名单，以控制迁移范围。

## Phase 1 Completion

已生成 [data-model.md](data-model.md)、三个 [会话与通信](contracts/document-session.md)/[原生工具](contracts/agent-tools.md)/[文件交换](contracts/python-exchange.md) 契约及 [quickstart.md](quickstart.md)。前后检查通过；无 before_plan/after_plan 钩子。下一步为 speckit-tasks；本轮未运行未来代码测试或性能测试，未生成实现代码。
