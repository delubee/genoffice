# Phase 0 Research: VS Code Office 编辑扩展

日期：2026-10-21。输入：[spec.md](spec.md)。以下是设计决策，不代表已通过原型或性能验收。依照 speckit-plan，已分别研究现有代码复用和官方 VS Code 接口。

## R1. 运行环境与交付边界

- **Decision**: 首版为桌面 workspace extension，`engines.vscode: ^1.109.5`，只使用该版本稳定 API；扩展宿主兼容 Node 22，开发环境沿用仓库 Node >=22.12/npm >=10。TypeScript 沿用仓库 5.9 系列。Python 使用工作区现有解释器，验收基线 Python 3.11+，不由插件安装。
- **Rationale**: VS Code 1.109 已正式提供 Agent Skills。1.109.5 源码使用 Electron 39.3.0，对应 Node 22.21.1；不能把独立 Electron 应用的版本和资源目录用于扩展宿主。
- **Alternatives considered**: Insiders/proposed API 会限制公开分发；浏览器扩展不能直接使用当前原生侧车，首版不选。
- **Sources**: [1.109 发布说明](https://code.visualstudio.com/updates/v1_109)、[1.109.5 package.json](https://raw.githubusercontent.com/microsoft/vscode/1.109.5/package.json)、[Electron 39.3.0](https://releases.electronjs.org/release/v39.3.0)。
- **Platform decision**: 首版验证 linux-x64 / Ubuntu 22.04、win32-x64 / Windows 11、darwin-arm64 / macOS 14 三种包；Remote SSH 先验证 Ubuntu 22.04 x64。其余架构、Alpine、纯浏览器和虚拟文件系统明确不声明支持。此为交付范围决定，不代表 VS Code 的能力上限。

## R2. 自定义编辑器与原生聊天工具

- **Decision**: `.xlsx` 注册 `CustomEditorProvider<WorkbookDocument>`；首版 `supportsMultipleEditorsPerDocument: false`。Webview 承接编辑，宿主管理每 URI 的文档、保存和工具。注册 `contributes.languageModelTools` 与 `vscode.lm.registerTool`，不实现另一套聊天、模型调用循环或 Python 执行服务。
- **Rationale**: 非文本文件需要自有文档模型；稳定扩展工具直接供原生 Agent 使用，并可访问打开的未保存状态。
- **Alternatives considered**: 只包装 CLI 修改磁盘不能读到当前未保存内容；仅预览不满足手工编辑；额外 MCP server 不是首版必要条件，避免双入口身份和权限混乱。
- **Sources**: [Custom Editors](https://code.visualstudio.com/api/extension-guides/custom-editors)、[LM Tools](https://code.visualstudio.com/api/extension-guides/ai/tools)、[固定版本 API](https://raw.githubusercontent.com/microsoft/vscode/1.109.5/src/vscode-dts/vscode.d.ts)。

## R3. 聊天目标必须显式绑定

- **Decision**: 提供“GenOffice: Bind Workbook Selection”命令，冻结当前文档、工作表、选区和版本，展示并可复制简短 binding 引用；也允许 context 工具使用明确 document URI 建立绑定。后续工具必须传 bindingId；变更另带 expectedRevision、transactionId。没有可靠目标时返回可选文档列表，不猜测最近焦点。
- **Rationale**: 没有稳定通用 API 可以监听所有原生聊天发送事件并截取发送瞬间的 Webview 选区。任务开始定义为成功建立 binding 的时刻，不等同用户按发送键。聊天、终端焦点变化不能更换已绑定对象。
- **Alternatives considered**: 使用 proposed chatContextProvider 不适合稳定版公开扩展；把 opaque toolInvocationToken 当持久聊天会话 ID 没有官方保证。
- **Sources**: [上下文提案](https://raw.githubusercontent.com/microsoft/vscode/main/src/vscode-dts/vscode.proposed.chatContextProvider.d.ts)、[官方上下文示例](https://raw.githubusercontent.com/microsoft/vscode-extension-samples/main/chat-context-sample/package.json)、[Proposed API 限制](https://code.visualstudio.com/api/advanced-topics/using-proposed-api)。

## R4. 共享代码与真实缺口

- **Decision**: 复用 `@genoffice/xlsx-gateway`、Rust xlsx-sidecar、同版本 Univer 和现有表格显示/保护/公式辅助模块。新增 `@genoffice/workbook-session`（Node 权威会话）、`@genoffice/sheets-editor`（浏览器编辑组件），抽取窄的数据源及宿主端口。现有 Electron 应用保留适配器；不复制大型 App.tsx，不仿造一个返回空值的完整 DesktopApi。
- **Rationale**: gateway 明确不依赖 Electron/Univer。`xlsx-sidecar-client.ts` 只依赖 Node 子进程与行协议，可移动后保留旧路径转导出；renderer 依赖全局 DesktopApi 的位置要改为注入。
- **Alternatives considered**: 整体嵌入现有 Electron App 会携带聊天、账户、主窗口、附件和导出耦合；重新写 Excel 引擎代价更高。
- **Evidence**: [gateway 包](../../packages/xlsx-gateway/package.json)、[sidecar client](../../apps/sheets/src/main/xlsx-sidecar-client.ts)、[renderer bridge](../../apps/sheets/src/renderer/mcp-bridge.ts)、[现有规划器](../../apps/sheets/src/renderer/plan-operations.ts)、[Rust Cargo.toml](../../apps/sheets/native/xlsx-engine/Cargo.toml)。
- **Gap**: `op-executor.ts` 可返回 partiallyApplied，并在执行后才报告 undoDropped；`save-actions.ts` 虽有普通保存 undo carry，但结构修改、两阶段保存及闭包式图表历史不能完整迁移。这些是必须补齐的功能，不是已具备的保证。见 [执行器](../../apps/sheets/src/renderer/op-executor.ts)、[保存](../../apps/sheets/src/renderer/save-actions.ts)、[undo carry](../../apps/sheets/src/renderer/undo-carry.ts)。

## R5. 单一权威状态与原子提交

- **Decision**: host 保存不可变打开快照 + typed overlay + 有序结构日志 + 稳定逻辑 sheet IDs + 单调 revision + 持久可序列化历史。Univer 是可重建投影及候选操作解释器，不拥有第二条用户历史。手工允许命令在执行前转为 semantic intent，进入与 AI/导入相同的每文档串行队列。
- **Transaction**: 预校验 → 在隔离的同版 Univer candidate 执行支持操作 → 提取全量受影响 changeset（值、格式、结构、图表、公式结果等）→ 检查版本/保护/撤销容量 → 写私有事务记录 → 原子移动 host 状态指针 → 发布视图和 CustomDocumentEditEvent。candidate 失败则丢弃；提交后可视投影失败则重建，不能再将已提交事务报成未提交。
- **Rationale**: 宿主在视图隐藏或重建时仍能准确读写；部分失败不污染主状态；undo 不依赖旧 renderer 闭包。
- **Alternatives considered**: 以 renderer 为唯一真相改动少，但失联与后台保存、历史恢复困难；把现有 AI planner 接成独立内存工作簿会形成分叉；只预校验而无隔离无法处理执行中失败。
- **Implementation validation**: offscreen candidate 的插件依赖、手工命令拦截、changeset 完整性和资源释放必须在实现早期原型中验证。未支持或不可序列化的命令默认禁止；不能降级成非原子直写。该验证属于已有明确决策的工程风险，不留未选择的架构分支。

## R6. 保存不更换当前会话

- **Decision**: 从最初不可变 base 和当前完整 overlay 生成 staged `.xlsx`，保真校验后替换目标。保存只移动 savedStateId 并更新 diskBaselineHash，不清空 overlay、不重开当前 sidecar/canonical session、不清空历史。undo/redo 增加 revision，但可回到原有 stateId；保存本身不增加内容 revision。
- **Rationale**: 将保存结果再当 base 后重放同一结构日志会重复插入行等；当前 undo-carry 也不满足完整跨保存历史。基于同一原始 base 保存不同历史状态可避免这两类问题。
- **Alternatives considered**: 保存后重新打开并改写低层 Univer IDs，不能可靠覆盖结构与图表闭包历史。
- **Detail**: 结构与新增图表等必须在私有临时包中完成全部 gateway 阶段，再一次性发布目标，不能将中间状态写到用户原件。saveAs 写新目标，遵守 VS Code 重新打开/关联文档生命周期，不自行修改 CustomDocument.uri；新实例生成新 epoch。实现 revert、backup、backupId 恢复；保存与历史移动共用队列。

## R7. Python 使用分型文件交换

- **Decision**: 版本化 manifest + typed NDJSON 行数据。导出由 host 合成 base+overlay，不从磁盘旧文件读。结果提供来源 snapshotId/哈希及目标信息；prepareImport 读取并复制到私有 staging、校验并返回不可变 importPlan，applyImport 只引用 planId 和 cursor。
- **Rationale**: 避免将 100k 数据经模型上下文搬运；防止准备和执行之间结果文件被改变；哈希验证完整性及来源关联，不声称证明 Python 的业务计算正确性。
- **Types**: 原始数值、numberFormat、公式及缓存计算状态分开；日期保留 Excel serial、1900/1904 系统，serial 60 不强行转 Gregorian 日期；文本“00123”、空串和空格与 null 区分。
- **Alternatives considered**: 纯 CSV 丢失类型及公式语义；openpyxl 直接保存打开的工作簿绕过当前会话；插件执行任意 Python 重复原生终端职责。
- **Limits decision**: 聊天读取 2,000 格；普通修改每批最多 2,000 展开单元格；导入默认分 1,000 格原子批；每个范围操作至多 200,000 格且仍需通过撤销预算。导出/结果文件每次最多 1,000,000 格且最多 64 MiB，达到任一上限拒绝、不截断。100k 输入分析不等于 100k 单批写回。整个多批导入不是整体原子，必须返回进度及可撤销已完成批次。
- **Existing limits**: 单个未压缩 OOXML 条目 500 MiB patch 限制继续保留；流式公式依赖固定预算 50,000 格继续生效；不绕过 existing gateway/schema/ZIP 限制。见 [DSL](../../packages/xlsx-gateway/src/domain/workbook-dsl.ts)、[保存网关](../../packages/xlsx-gateway/src/gateway/xlsx-package-io.ts)、[公式预算](../../apps/sheets/src/renderer/univer-state.ts)。新文件交换和历史预算是首版工程边界，须实测。

## R8. 原生保存/撤销与持久历史

- **Decision**: 每个成功语义事务仅发一个 CustomDocumentEditEvent；undo/redo 回调移动 host timeline。Webview 快捷键转发到相同历史，不并行启用 Univer 用户历史。预检每事务序列化双向变化 <=64 MiB，当前 retained timeline <=256 MiB，达限在执行前拒绝并说明关闭重开可释放历史；保存不自动剪裁。
- **Rationale**: 满足执行前可撤销保证，避免大操作做完才丢弃 undo；host 引用文件化的 chunks，而不是把工作簿整体复制到每个历史项。
- **Recovery**: 使用扩展私有会话目录和 VS Code backup 生命周期；完整 manifest 写入成功才确认副本，异常恢复仅承诺最近已确认备份。恢复有新 sessionId，旧 binding/snapshot 写入资格失效。
- **Sources**: [Custom Editor API](https://code.visualstudio.com/api/extension-guides/custom-editors)、[固定版本 API](https://raw.githubusercontent.com/microsoft/vscode/1.109.5/src/vscode-dts/vscode.d.ts)。

## R9. 技能安装与 Remote SSH

- **Decision**: 随扩展打包 `genoffice-workbook` 技能和引用资料，首版用显式命令安装到用户选择的 workspace `.github/skills/genoffice-workbook/`，已有文件先显示差异并由用户选择，不自动覆盖。技能只指导绑定、工具、Python 文件交换和结果检查。
- **Rationale**: 稳定 `contributes.chatSkills` 已存在，但 1.109 发布说明目录示例与现行文档 SKILL.md 路径示例有差异；首版用已支持的 workspace 发现路径，避免未验证的打包行为。不是声称“无稳定技能贡献接口”。
- **Alternatives considered**: 发布扩展技能贡献可在固定最低版本上验证后另行启用；不默认往用户全局配置写文件。
- **Remote decision**: `extensionKind: ["workspace"]`，sidecar、快照、Python 工作文件与原生终端同在 workspace host，Webview 只在客户端渲染。解析 URI、验证 scheme/authority 并确认宿主本地映射后才取文件路径。首版拒绝不能映射到宿主文件系统的提供器。
- **Sources**: [Agent Skills](https://code.visualstudio.com/docs/agent-customization/agent-skills)、[Extension Host](https://code.visualstudio.com/api/advanced-topics/extension-host)、[Remote Extensions](https://code.visualstudio.com/api/advanced-topics/remote-extensions)。

## R10. 验证、发布与治理

- **Decision**: 现有 Vitest/Rust 测试复用，新会话/契约加 Vitest；新扩展使用稳定 VS Code 集成测试宿主，浏览器编辑组件用 Playwright；三平台 VSIX 分别带独立 Rust sidecar，无 Electron 原生模块 ABI 依赖。Rust edition 2024，沿用 Cargo.lock，构建流水线记录通过的 rustc 版本，执行 Windows CRT 静态链接配置。
- **Rationale**: 浏览器 UI、Node 会话和二进制三个边界需要独立验证；模型调用具有不确定性，确定性工具测试与人工真实聊天验收分别记录，不能把模拟工具成功当成模型保证。
- **Baseline**: Ubuntu 22.04 x64、4 vCPU、8 GiB RAM、SSD，固定 VS Code 1.109.5，无其他扩展；20 次打开/修改记录原始时延，至少 19 次满足规格门槛；聊天手测环境另记模型与账号，不混入本地时延。
- **Governance**: 宪章是未填模板，没有可执行原则；不能从模板示例虚构阻断或批准。按现有 CLAUDE.md 的主题/i18n/构建规则设计。前后规划检查均记录在 plan.md；没有扩展钩子配置。

## Research disposition

已对目标版本、技能接入、选区绑定、权威状态、原子性、保存历史、文件交换、平台及验证作出明确决定，没有遗留设计澄清标记。文档只确定方案；原型、容量测量、三平台打包和端到端聊天仍是后续实现验证工作，不能视为已完成。
