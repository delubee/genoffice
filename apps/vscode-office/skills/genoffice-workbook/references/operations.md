# 原生工具与操作 v1

输入包含 `schemaVersion: 1`。成功返回 status、sessionId、revision、summary、data；失败返回稳定 code、details、retryable。
先检查 status，再解释 data。

| 工具 | 输入与行为 |
| --- | --- |
| genoffice_workbook_context | bindingId 或 documentUri，可附 sheetId/range；返回身份、版本、工作表、冻结选区、保护信息、能力和限制；无目标仅列选项 |
| genoffice_workbook_read | bindingId、sheetId、range、expectedRevision?、includeFormats?；读取当前值、显示值、公式和计算状态 |
| genoffice_workbook_aggregate | bindingId、sheetId、range、expectedRevision?；只读统计，不写辅助格 |
| genoffice_workbook_apply | bindingId、expectedRevision、transactionId、summary、operations、usesFrozenSelection?；一批原子修改、一个撤销步骤，不自动保存 |
| genoffice_workbook_operationStatus | sessionId + transactionId 或 sessionId + planId；查询事实，不自动重新打开关闭的文件 |

read 普通最多 2,000 格，带格式最多 200 格；返回体超过 1 MiB 会拒绝，要求缩小范围，不截断。
aggregate 最多 1,000,000 格，最多 50 个高频值；只对 number 类型求和，不混入日期或数值外观文本。
空字符串是非空文本，blank 才是空白。未验证公式缓存会使统计明确失败。

## 操作子集

始终以 capabilities 为准。首版支持：

- set_cell、set_formula、clear_cell、set_range、clear_range；
- format_range：常用字体、颜色、对齐、换行、边框、数字格式；
- copy_range、fill_range：复制值/公式并保留目标格式；过滤复制不支持；
- add_sheet、rename_sheet、delete_sheet，受限行列插入/删除；
- add_chart、edit_chart、delete_visual：本会话新增的 column/bar/line/pie 基础图表。

原文件图表及高级对象不按可编辑对象推断。保护、合并、表格、命名区域等可能限制结构变换；
不能保真的操作会在提交前拒绝。图表和结构操作另受引用材料与公式依赖预算限制。

set_cell 的字符串为文本；set_formula 显式写公式。set_range 沿用既有 DSL 的公式前缀规则；
需要以 `=` 开头的字面文本时使用 set_cell 或 typed import。fill/copy 会移动相对公式引用。

```json
{
  "schemaVersion":1,
  "bindingId":"context返回的绑定",
  "expectedRevision":7,
  "transactionId":"唯一事务标识",
  "summary":"更新销售额",
  "operations":[{"op":"set_cell","sheetId":"logical-sheet-id","address":"D2","value":120}]
}
```

每批最多 1,000 个 DSL 操作，普通修改最多展开为 2,000 格；范围操作最多 200,000 格。
实际双向变化超过 64 MiB 或保留历史超过 256 MiB 会拒绝；保存不会清历史腾预算。
相同的内容寻址存储块只计一次。

结构与内容分批：先提交结构，重新读取/绑定，再写内容。
usesFrozenSelection 默认 true；结构后的原选区不能仅靠刷新 revision 恢复有效。
明确新区域需重新读取后设 false，或创建新绑定。

receipt 包含新 revision、stateId、affectedTargets 和 calculationState。
undo/redo 恢复旧内容根但 revision 继续增长；数据看起来相同也不能复用旧 revision。
