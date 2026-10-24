# Data Model: Workbook Session

设计状态，关联 [spec.md](spec.md)、[plan.md](plan.md) 及 [会话契约](contracts/document-session.md)。字段是规范语义，不是最终实现类声明。

## Identity and versions

- `documentUri`：完整 VS Code URI，包含 scheme/authority/path；不能用 basename 或客户端文件路径当身份。alias/realpath 检测防止同一物理目标并行写入。
- `sessionId`：每次打开生成的随机身份；dispose、revert 或恢复新会话后，旧绑定与计划失效。
- `revision`：会话内单调非负整数；内容 commit/undo/redo 每次 +1；save、选区变化、索引进度不改变内容 revision。
- `stateId`：不可变语义状态 ID；undo 回到既有 stateId 但获得新 revision，避免 ABA。`savedStateId` 标记已保存语义状态，决定 dirty；VS Code 的 edit event/save 生命周期同时维护界面标记。
- `logicalSheetId`：会话内稳定的工作表身份；与 OOXML part/rId、Univer unit/style IDs 解耦，保存后不变化。名字唯一且符合现有 DSL 校验；sheet 删除保留历史中的映射供撤销。

## Entities

### WorkbookDocument

| 字段 | 含义/约束 |
| --- | --- |
| documentUri / sessionId / revision | 权威身份与版本 |
| originalBase | 不可变原包私有副本、SHA-256、archive manifest、sidecar session；保存不替换它 |
| stateId / savedStateId | 当前/已保存状态，值相等时语义 clean |
| sheets | logical ID、名字、part映射、数据范围、保护/只读能力 |
| overlayRoot | typed cell/format/feature 状态与 ordered structural log；永远相对 originalBase |
| diskBaselineHash | 最近成功读取/保存的目标内容摘要，区别于 originalBase hash |
| timeline / receipts | 历史根与幂等提交收据 |
| lifecycle | opening / ready / preparing / committing / saving / conflict / recovering / disposed |

唯一所有者是 workspace extension host；投影、Agent 与 Python 都不能另建独立权威工作簿。base/overlay/timeline 被引用期间不可清理；关闭后仍被 backup 引用的文件保留。

### WorkbookState and CellRecord

WorkbookState 由 immutable overlay root、ordered structural log、sheet mapping、chart/table/feature 状态组成。仅需保存相对 base 的变更，但导出/读取必须合并完整当前语义。

CellRecord 包含 logicalSheetId、A1 地址、typed rawValue、formula（可缺省）、calculationState、cachedValue、numberFormat、受支持 style patch。文本空串与 null 不相等；`=...` 字面文本不能按前缀自动解释为公式。日期保留 serial 与 dateSystem，百分比保留比例 rawValue，显示文本仅作辅助。

结构操作采用明确前后坐标空间；结构提交使旧 revision 的地址全部失效。候选生成式依赖、图表数据缓存、sheet 名引用重写也必须进入结果，不能只持久化显式单元格编辑。

### TaskBinding

字段：bindingId、sessionId、documentUri、logicalSheetId、frozenRange（可空）、boundRevision、createdAt、status。

创建来自显式用户绑定命令，或带明确 URI 的 context 工具。status 为 active / stale / invalidated；修改后 stale 的 selection token 可用于说明目标，但不授权修改，必须读取当前版本。close/revert/recovery 为 invalidated，禁止自动绑定到同路径新会话。绑定不是聊天身份，也不能跨窗口复用。

### Transaction and Receipt

字段：transactionId、bindingId、expectedRevision、origin（manual / agent / import / history）、label、normalizedIntent、payloadHash、beforeStateId、afterStateId、forward/inverse chunk refs、serializedBytes、status、commitRevision。

状态：received → preparing → prepared → committed；提交前任意失败 → rejected / cancelled。只有成功持久 commit record 并移动状态指针才算 committed。提交后的视图故障为 projectionPending，而非 rejected。重复 transactionId+同 payload 返回原收据；不同 payload 为 idempotency_conflict；幂等检查先于过期 revision 检查，确保网络丢响应后的重试可取回成功状态。

收据含 sessionId、transactionId、status、revision、stateId、affectedTargets、summary、calculationState；无效输入不形成已应用收据。持久收据与会话生命周期共存，恢复到新 session 后不得用于向新会话重试旧修改。

### HistoryEntry and RecoveryPoint

HistoryEntry：entryId、before/afterStateId、origin、label、forward/inverse refs。host 持有，VS Code undo/redo 回调引用 entryId；不序列化 JS closures 或 Univer 内部 unit IDs。新内容 commit 清空 redo 分支；save 不改变 timeline。每事务双向变化 <=64 MiB，保留 timeline <=256 MiB；超出拒绝新修改，不在保存时静默删历史。

RecoveryPoint：backupId、originalUri、createdAt、baseHash、overlayRoot、stateId、manifestHash、完整性状态。先写所有引用文件，再原子发布 manifest，最后确认 backupId。恢复至少保留该副本内容，未保存恢复标记为 dirty；首版无需跨进程恢复整个 undo timeline，旧 session 失效。

### DataSnapshot

字段：snapshotId、sessionId、sourceStateId、sourceRevision、logicalSheetId、range、schemaVersion、dateSystem、dataUri、dataSha256、manifestSha256、rowCount、columnCount、cellCount、complete、calculationState。

sourceRevision 在导出过程中固定；使用 immutable read view，后续编辑不污染导出。快照仅在数据完整且已验证持久化后 complete=true。host 私有 registry 保留源元数据，不信任 Python 可编辑 manifest 中的版本声明。

### PreparedImport and ImportProgress

字段：planId、snapshotId、bindingId、initialExpectedRevision、privateResultHash、target、mode、batchCount、nextCursor、chainExpectedRevision、completedReceipts、status、expiresAt。有效期 30 分钟；使用后的 batch tokens 不重放。

状态：prepared → applying → completed；错误/取消 → failed / cancelled / stale。初始 snapshotRevision 必须等于当前 revision。后续仅允许该计划前一成功批次生成的 revision 继续；任何其他内容 commit 都使计划 stale。分批默认 <=1,000 个普通单元格，每批原子且一个历史项，整份文件不承诺原子。新工作表创建及首批内容在同一隔离候选中分子阶段验证后一次提交；失败不遗留空表。之后按返回的稳定 logicalSheetId 继续。

### SkillInstallation

字段：workspaceUri、destination、bundledVersion、installedDigest、userModified。目标 `.github/skills/genoffice-workbook/`。已存在用户修改时需差异及明确选择，拒绝更新不影响手工编辑或已有工具。

## State transitions

```text
open -> snapshot/index -> ready(state0, revision0)
ready -> prepare(candidate) -> commit(stateN, revision+1) -> ready
prepare fails/cancelled -> ready(original state, same revision)
ready -> undo/redo -> ready(existing stateId, revision+1)
ready -> save -> ready(same state/revision; saved marker + disk hash updated)
save fails -> ready(dirty retained) or conflict
external disk drift -> conflict -> saveAs/cancel OR confirmed revert(new session)
any -> dispose -> old bindings/imports invalid
backup -> confirmed recovery point -> restore(new session, dirty if unsaved)
```

## Resource and validation rules

协议及文件 `schemaVersion=1`，未知版本拒绝。事务按文档串行，索引与只读合成可并行但绑定 immutable stateId；candidate 完成后仍重新检查 revision。取消在 commit point 前丢弃候选，之后返回已提交收据。消息、导出及历史的容量分别计量，不能只按单元格数推断内存占用。

所有路径校验 scheme/authority、workspace 边界和真实映射；Python 只获得用户可见交换文件，不获得私有 canonical 目录写权限。manifest 哈希证明输入文件关联及完整性，不能证明外部程序业务逻辑正确。业务正确性通过验收数据的独立 expected results 验证。
