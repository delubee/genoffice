# Contract: Document Session and Editor Bridge v1

关联：[data-model](../data-model.md)、[Agent tools](agent-tools.md)。规范性 MUST 为实现要求，示意消息不是实现代码。

## Host service boundary

会话服务是 Node 库，不导入 vscode、Electron、React 或 Univer。依赖窄端口：

| 端口 | 操作/责任 |
| --- | --- |
| WorkbookSource | openSnapshot、readBaseRange、readFeatures、archiveManifest、close；底层为已校验 sidecar client |
| CandidateInterpreter | prepare(current state, intent, cancel) → serializable changeset；隔离同版 Univer，不改变可视或 canonical 状态 |
| ResourceStore | 私有 snapshot/chunk/backup，workspace 导出路径，宿主本地 URI 映射 |
| PackageWriter | 从 original base + current overlay 生成 staged 包，校验 preservation，受保护的 publish |
| HostLifecycle | edit event、保存完成、提示/冲突、投影广播；由 VS Code 或 Electron 适配 |

公共服务操作：open、describe、bind、read、prepareTransaction、commit、undo、redo、save、saveAs、revert、backup、exportSnapshot、prepareImport、applyImportBatch、dispose。只有服务可移动 revision/state/history。允许纯只读消费者使用冻结 read view；不能由 CLI 独立覆盖同一打开文档。

## Message framing

Webview 使用 acquireVsCodeApi().postMessage，host 使用 webview.postMessage。每个消息包含 `protocolVersion:1`、`requestId`、`sessionId`、`method`、`payload`；请求成功为 `{requestId,sessionId,ok:true,result}`，失败为 `{requestId,sessionId,ok:false,error:{code,message,details,retryable}}`。拒绝未知字段/方法与错误 session；不接受来自 Webview 的任意文件路径、shell 或脚本执行命令。

消息体 UTF-8 JSON 上限 1 MiB，大内容分 chunk（chunk <=256 KiB），含 transferId/index/count/digest；收到全部且校验正确才用于事务，不边收边改。单次准备候选及 Node↔sidecar 的超时按具体命令处理；通用控制请求 30 秒，文件生成 180 秒。超时后 caller 使用 transactionId 查询结果，不能把 timeout 自动判成未提交。

| 方法 | 方向 | 核心语义 |
| --- | --- | --- |
| view.ready | view → host | 报告协议、支持 capabilities；返回当前 session/revision/state 和必要首屏 |
| view.selection | view → host | sheet/range/editorStatus，附当前 projectionRevision；非内容修改 |
| view.intent | view → host | 手工 command 转 semantic intent，expectedRevision/transactionId；未批准不修改主投影 |
| candidate.prepare | host → view | 固定 stateId 和明确 intent，在独立 runtime 中计算 changeset |
| candidate.result | view → host | changeset refs、基线、错误或取消；host 不盲信版本/字节数/能力声明；state/sheet ID由host预分配，host对同一material独立重放纯数据语义并对比完整结果 |
| state.committed | host → view | receipt、new revision/state、可用投影变化；失败则请求完整重建 |
| view.resync | 双向 | 按 host state 重建；不得用可视状态覆盖 canonical |
| history.request | view → host | 请求统一 undo/redo，经 VS Code 历史入口调用对应回调 |
| source.read | view → host | session/logical sheet/区域/state，映射 originalBase+overlay；不直连全局 desktopApi |
| view.dispose | view → host | 释放 projection/candidate；是否关闭文档由 CustomDocument 生命周期决定 |

只支持已注册 semantic intents；加载/索引/主题和计算缓存的内部命令带 internal origin，不发用户 edit event。输入法合成期间 read/export 要求 flush 当前编辑，不能完成则 editor_busy，不读到一半输入。

## Transaction semantics

1. 在每文档队列中校验 session/binding、idempotency、capabilities、expectedRevision、保护、操作 schema 和范围。
2. 创建 immutable read view；在独立 candidate 运行全部受支持子阶段，收集由该操作产生的值/格式/结构/图表/公式/保护变化。可使用现有 executor 作为 candidate 内部组件，但任何 partiallyApplied 只影响候选，必须丢弃该候选。
3. host 校验 changeset 引用和状态、计算真实序列化大小、历史容量、取消及当前 revision；不完整逆变更禁止提交。
4. 将 chunks 和含 receipt 的 commit manifest 写私有目录，全部成功后原子发布新 canonical root。写入失败不会移动根；无需因 projection 成败回滚一个已经提交的 canonical 事务。
5. 发一个 CustomDocumentEditEvent 和 state.committed。网络丢响应/投影失败以 receipt 为准，允许查询及重建。崩溃点若 commit manifest 已发布，恢复必须识别事务已提交。

原子单位为一个 transaction；多批文件导入/任务不是一个整体事务。prepared 期间其他修改只可等待或先完成导致候选过期；不能无检查重放。每事务执行前确认双向记录 <=64 MiB、整个 retained timeline <=256 MiB，否则 undo_budget_exceeded。

## CustomEditor lifecycle

- `openCustomDocument`：URI 映射/alias 锁、workspace trust、文件格式与完整性检查；有 backupId 时校验并恢复该备份；创建新 sessionId。无效文件不创建可保存空文档。
- `resolveCustomEditor`：加载本地构建 Webview；nonce CSP、asWebviewUri 与 localResourceRoots；限定首版一文档一编辑视图；retainContextWhenHidden 用于保持普通后台候选，内存退出可重建。
- `onDidChangeCustomDocument`：每个成功事务一次 edit event，label、undo/redo 回调引用稳定 HistoryEntry。undo/redo 本身不再重复发新 edit event；revision 仍递增。
- `saveCustomDocument`：等待当前编辑 flush；冻结 state；检查磁盘 baseline；私有文件完成所有 gateway 子阶段及保真校验；安全替换；更新 savedStateId/hash，不清 journal/history、不换 session。
- `saveCustomDocumentAs`：同样生成完整包并检查目标覆盖授权；源文件不自动改写。由 VS Code 决定目标如何重新打开，provider 不自行修改 uri；若创建新 CustomDocument，使用新 sessionId，旧 token 不迁移。
- `revertCustomDocument`：遵守明确 reload/放弃选择，重新载入磁盘；旧 binding/import 失效，清理历史仅在该明确 revert/新会话边界执行。
- `backupCustomDocument`：备份 base+完整 overlay 和必要 feature maps，manifest 原子落盘后返回 backup id 与 delete；恢复副本引用未清除前 base 不可删除。
- `dispose`：使 binding/import 失效，停止候选与 sidecar 引用；根据 backup refs 清理，不能影响其他文档。

## Save conflicts and failure handling

保存队列串行；目标当前摘要不等于 baseline 则进入 external_change_conflict。写临时包、执行重排、更新图表等都不能提前写用户文件。目标目录只读、空间不足或用户取消保留 canonical/history 和 dirty。

对外部编辑器，文件系统不提供通用跨进程 compare-and-swap：发布前后检查及 watcher 是检测，不构成绝对排他。冲突时保留足够的 old/staged 副本供另存为或重载；不得将测试中的冲突检测外推为与任何第三方程序的无条件原子合并。

## Stable failures

`invalid_document`、`unsupported_format`、`unsupported_platform`、`untrusted_workspace`、`unsupported_operation`、`protected_range`、`editor_busy`、`view_unavailable`、`revision_conflict`、`session_closed`、`external_change_conflict`、`undo_budget_exceeded`、`range_limit_exceeded`、`formula_budget_exceeded`、`idempotency_conflict`、`operation_cancelled`、`save_failed`、`recovery_invalid`。

error details 至少给明确目标、expected/current revision（适用时）、是否已提交、可重试方式；文档数据只作为数据，不作为可执行指令。只读工具不声称修改成功，失败不得静默退回磁盘旧内容。
