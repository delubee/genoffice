# Quickstart Validation: VS Code Workbook Editor

状态：六个故事的代码和本地确定性验证已实现，证据见 `apps/vscode-office/test-results/`。Linux VSIX 已生成；真实模型聊天、Windows/macOS/SSH 和指定参考设备验收仍待完成。本文未测环境不代表已通过。

契约参考：[会话](contracts/document-session.md)、[原生工具](contracts/agent-tools.md)、[Python交换](contracts/python-exchange.md)、[数据模型](data-model.md)。

## Prerequisites

- 开发环境 Node >=22.12、npm >=10；Rust 工具链能够构建现有 edition 2024/Cargo.lock，并记录实际rustc版本。开发需要cargo，安装好的VSIX不需要。
- VS Code 1.109.5 及一个当前稳定版本作为独立验证环境；桌面运行、可信的真实文件工作区。
- 程序分析使用工作区 Python 3.11+；首个汇总场景仅用标准库 json/decimal，避免把依赖安装当成插件功能。另测试无Python及缺库错误。
- 真实聊天检查需要原生Agent、可用模型及启用扩展文档工具；确定性工具/保存测试不需要模型账户。
- 参考性能环境：Ubuntu22.04 x64，4 vCPU、8GiB RAM、SSD，本地workspace，扩展以构建后版本运行，仅启用被测扩展。Windows11 x64、macOS14 arm64及Remote SSH Ubuntu22.04 x64分别记录额外结果，不混合采样。

## Build and launch

现有开发入口（从仓库根运行）：

```bash
npm ci
npm run native:build -w apps/vscode-office
npm run typecheck -w @genoffice/xlsx-gateway
```

build/typecheck/test/fixtures/package 入口已实现；发布验收状态详见 apps/vscode-office/IMPLEMENTATION-STATUS.md。夹具生成到 apps/vscode-office/test-fixtures/generated/：

```bash
npm run build -w @genoffice/workbook-session
npm run build -w @genoffice/sheets-editor
npm run build -w apps/vscode-office
npm run fixtures -w apps/vscode-office
code --extensionDevelopmentPath="$PWD/apps/vscode-office" "$PWD/apps/vscode-office/test-workspace"
```

两个共享包使用上述scope名，扩展使用合法manifest名称genoffice-workbook并以workspace路径调用，已提供build/test/typecheck等脚本；fixture脚本创建固定样本及独立预期结果，不自动安装技能、Python或修改用户模型设置。Linux无显示的集成测试使用xvfb-run；真实聊天人工测试在可交互桌面完成。

## Deterministic fixtures

`npm run fixtures -w apps/vscode-office` 已生成并记录以下摘要：

- `sales-100k.xlsx`：Data表含表头共10,000行、10列，四个地区、数值销售额、日期/百分比/文本编号；另一个Config表含受支持基础公式与格式。固定seed和独立`expected.json`，规定Data!D2原值100、未保存修改为120，属于North，其他输入固定。
- `types-1900.xlsx`、`types-1904.xlsx`：包含serial60、日内小数、00123、空文本、blank、boolean、百分比和公式ready/pending/error。
- `preservation.xlsx`：含原始图表、未知可保留OOXML part、受保护区域和合并；检查未触及条目内容/CRC不变。
- 同名不同目录工作簿、损坏文件、外部覆盖副本、低历史预算测试配置及导入文件篡改样本。

已有 `apps/sheets/scripts/generate-large-xlsx.ts` 可生成单表100k规模探针，但它不是上述双表销售/类型验收夹具，不拿它代替SC-001/004。

## V1 — Manual editing and history

1. 不登录或启用AI，在Explorer对sales-100k.xlsx使用“Open With → GenOffice Workbook”。切换Data/Config，核对值、公式和常用格式。
2. 将Data!D2由100改为120，不保存，检查dirty；新增公式、改number format，核对结果。
3. 按手工、一次工具批次、手工顺序修改；从编辑器快捷键及VS Code命令分别撤销/重做，逐次对照expected state。
4. 保存，继续撤销到保存前状态、重做并再次保存重开，确保历史未被save清空，结构/图表修改也能跨保存撤销。
5. 另存为新目标，原文件保持先前已保存状态；关闭dirty文档分别选择取消、放弃、保存。

期望：FR-001–003、021–023，SC-002/007全部满足；候选执行器/渲染库无额外用户历史导致双撤销。

## V2 — Native chat and skills

1. 执行“GenOffice: Install or Update Workbook Skill”，选择该workspace；核对生成`.github/skills/genoffice-workbook/SKILL.md`和资料。修改技能，再次安装时选择保留，确保内容未覆盖。
2. 在Data选D2:D10，执行“GenOffice: Bind Workbook Selection for Chat”，把返回binding引用粘贴到原生聊天。此时任务目标才算建立。
3. 请求“解释该区域，不修改”，核对AI读到未保存D2=120，dirty及内容不因查询而变化。
4. 请求“写入本列合计公式并设置百分比列格式”；核对实际单元格、公式结果、批次摘要和Undo。技能不要求另配文档AI密钥。
5. 在已绑定后切换同名文件/工作表/聊天/终端，确保已有binding仍指向原目标。无binding仅说“这一列”应要求明确，而非使用最近焦点猜测。

期望：US2/US6，SC-003/008满足。记录使用的Agent、模型、工具启用状态及技能版本；不把模型未调用工具归因为文档已成功修改。

## V3 — Python end-to-end

1. 重开未修改夹具，手工将Data!D2改为120且不保存，明确绑定Data。
2. 原生聊天要求：“用Python读取当前工作簿的数据，按地区汇总销售额，写入新的Summary工作表并创建柱状图；保留源数据。”
3. 检查执行顺序：context/export → workspace脚本 → result manifest → prepareImport → 一批或多批applyImport → 最新context/read → chart apply。
4. 查看`.genoffice/work/<snapshotId>/`中的输入、脚本和结果；完整100k数据不在聊天正文中逐条出现。程序输入D2=120，North结果较原始fixture正确值增加20，其他地区不变；所有结果对照expected.json。
5. 继续手工编辑Summary，撤销/重做、保存重开，确认图表数据引用和源数据完整。
6. 对types夹具执行值模式/显式公式模式往返，逐项检查百分比、日期系统和文本编号；空结果不清空区域。

期望：SC-004及FR-011–016；源数据变更仅有明确的手工D2修改。脚本失败或缺依赖时无不完整导入。

## V4 — Imports, cancellation and budgets

1. 创建100k格合法结果文件，prepare后按1k格批次回写，核对batchCount、cursor、receipt；单批撤销不应撤销其他批。
2. 在第3批完成后取消：前3批保留且可撤销，剩余不写入，报告为cancelled而非completed。
3. 执行中注入candidate内部第N项失败：该批任何变化均不可见，前批不受影响；UI投影提交后故障则按已提交receipt重建。
4. prepare后修改公开result文件：已准备私有副本不受影响；重新prepare被改坏的文件则hash/type校验失败。更改sourceRevision不能绕过host registry。
5. 在批间手工修改、undo或另一个AI任务提交，下一批必须stale；不得仅刷新revision继续旧结果。
6. 超过2k展开写、200k范围、1m/64MiB交换、64MiB单事务或256MiB历史预算，执行前明确拒绝；目标不变。大字符串与同样格数的小数值分别测量。

期望：每批原子、完整可撤销、准确部分进度；100k输入汇总与100k全量回写分别测试，不能混为同一容量保证。

## V5 — Conflict and failure injection

1. AI读后手工编辑，回写旧revision失败；插行后旧地址也失败。文档关闭重开后旧binding不能用，即使URI相同。
2. 存在dirty时从外部改磁盘文件，save报conflict，保留双方并可另存为。对发布前后外部写入窗口做故障注入，保留必要副本并如实报告；不声称任意第三方写进程绝对互斥。
3. 提交成功后丢弃工具响应，重试同transactionId得到原receipt；同ID不同payload被拒绝，sheet/chart不重复增加。
4. 注入磁盘满、目标只读、staged包写失败和preservation失败，原件、dirty内容和历史均保留。
5. 确认backup成功后终止测试Extension Host，从backupId恢复，保留已确认副本，标记恢复时间；旧session/import失效。
6. 隐藏或重建Webview，读取/导出/保存从host当前状态完成；候选解释器不可用时修改明确失败，不读取旧磁盘状态代替。

期望：SC-003/006/007及FR-017–028；预提交失败没有部分落地，提交后故障不能错误报告“未改动”。

## V6 — Packaging, platforms and permissions

按实现阶段约定提供如下命令：

```bash
npm run test -w @genoffice/workbook-session
npm run test -w @genoffice/sheets-editor
npm run test:integration -w apps/vscode-office
npm run test:e2e -w apps/vscode-office
npm run benchmark -w apps/vscode-office
npm run package -w apps/vscode-office -- --target linux-x64
```

发布流水线分别在对应宿主构建/测试win32-x64、linux-x64和darwin-arm64二进制及VSIX；最低VSCode1.109.5和当前稳定版都测试。安装VSIX到干净profile后不装cargo/GenOffice独立应用，手工编辑仍可用。资源CSP、workspace trust与schema拒绝测试不跳过。

Remote SSH测试连接Ubuntu22.04 x64：确认扩展/sidecar/Python/交换目录都在远端，UI在客户端；不使用客户端路径。禁用Agent工具检查明确启用提示；拒绝原生执行确认检查没有备用Python执行器绕过；无Python环境保持手工编辑。

回归采用受影响现有workspace的typecheck/Vitest及Rust原生测试，并运行现有主题/i18n/许可门槛；命令从仓库真实package脚本选择，不用未来新包命令替代已有回归。完整根测试只在共享层改动风险或发布门槛需要时执行。

## Performance and acceptance record

对SC-001和SC-005各采20次：记录设备、OS、VSCode/扩展/sidecar版本、fixture SHA、冷暖启动、完整原始样本。打开从用户打开动作至首屏可编辑，不用sidecar open耗时替代；修改从校验完成至UI显示完整结果，公式计算/模型等待单独记录。若少于19次达到阈值，不通过；应改进实现或提出规格调整，不直接修改目标掩盖失败。

验收报告分别列出SC-001–008及V1–V6，附确定性测试与真实聊天证据。当前规划阶段所有实施验证均为“未执行”，不能把本文的步骤当成结果。


## Implemented verification commands

```bash
npm run vscode:build
npm run vscode:typecheck
npm run vscode:test
npm run test:e2e -w @genoffice/sheets-editor
npm run test:integration -w apps/vscode-office -- --label minimum --fail-zero
npm run test:capacity -w apps/vscode-office
npm run benchmark -w apps/vscode-office
npm run notices
npm run package -w apps/vscode-office
npm run test:package -w apps/vscode-office
```

无显示 Linux 的 GUI 命令需 xvfb-run；Python capacity 为约 11 分钟的独立容量/正确性测试。benchmark 的本机结果不替代 4vCPU/8GiB/SSD 参考环境。安装命令只生成或验证本地 VSIX，不发布。
