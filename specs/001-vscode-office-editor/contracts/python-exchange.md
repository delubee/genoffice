# Contract: Typed Workbook Exchange v1

用于原生 Agent 的终端/Python 与当前文档之间交换数据，不内置 Python 执行器。关联 [data-model](../data-model.md)、[tools](agent-tools.md)。

## Directory, ownership and limits

默认目录为所选 workspace `.genoffice/work/<snapshotId>/`：input.manifest.json、input.ndjson、analysis.py（Agent创建）、result.manifest.json、result.ndjson。用户可查看这些文件；扩展私有 session/chunk/history 目录不向程序开放写接口。多 root 必须选定一个 root；远程时文件和 Python 均位于 workspace host，工具同时返回 URI 与经过验证的宿主路径。

每份输入/结果最多 1,000,000 格、UTF-8 数据文件最多 64 MiB，任一超过即拒绝并建议范围拆分，不能截断。每条 NDJSON 记录不超过256 KiB，过宽行可分连续 segment，禁止将一行限制误报成总列上限。manifest <=256 KiB。导出逐块读取当前 immutable view，含 overlay 和坐标映射，不对整个文件执行 JSON.stringify。mtime 不是身份/完整性依据。

## Input manifest

必填字段：

| 字段 | 约束 |
| --- | --- |
| schemaVersion | 整数1 |
| kind | workbook-snapshot |
| snapshotId / sessionId / documentUri | 宿主生成的源身份 |
| sourceRevision / sourceStateId | 导出时内容版本与状态，整个导出不变 |
| sheet | logicalSheetId、name |
| range | 明确A1矩形，与shape一致 |
| shape | rows、columns、cellCount，rows*columns=cellCount |
| dateSystem | 1900 或1904 |
| calculationState | ready/pending/error；pending/error不是有效0 |
| data | 相对文件名、sha256、bytes、recordCount、encoding=utf-8、format=typed-ndjson-v1 |
| complete | true 才可供已完成分析使用 |

manifest 最后原子发布；数据落盘、计数和哈希校验成功前不能 complete=true。源 registry 在私有宿主中保存同一 manifest 摘要，Python 更改公开 sourceRevision 不能获取新版本授权。读取过期快照仍可用于用户查看，但不得回写原会话。

## Data records and types

每行数据是一条 `{rowOffset, columnOffset, cells:[TypedCell,...]}`，偏移从0开始且相对 range 起点；按行/列递增，可用多个 segment 覆盖一行。必须恰好覆盖整个矩形，不重叠、不缺格、不超界；blank 也显式记录，以便保留空白含义。reader 不依赖 NDJSON 物理行号推断工作表行号。

| TypedCell.kind | 字段/语义 |
| --- | --- |
| blank | value=null，区别于空文本 |
| string | value为原文本，含空串、换行及00123；即使以=开头也仍为文本 |
| boolean | value为true/false |
| number | value为有限数值，可附numberFormat；百分比0.25+0%不变成25 |
| date | serial为有限数值、dateSystem、numberFormat，保留日内小数，不自动加本地时区 |
| error | excelCode，如#REF!；不当作普通数字或空值 |
| formula | formula以=开头、cached为非formula TypedCell或null、calculationState；缓存与公式独立 |

可附 displayText 供人类查看，但程序不能据其替代 raw value。1900系统 serial60 保留为 Excel 特殊日期序数，不生成假装合法的 Gregorian ISO 时间。1904值不按1900系统解释。数字遵循工作簿数值精度；超过该精度的编号须为string，不悄悄转数值。

## Result manifest

必填 schemaVersion=1、kind=workbook-result、source（snapshotId、sourceManifestSha256、sourceDataSha256、sourceRevision）、output（同data元信息）、shape、target、mode、complete=true。target只能为 `{sheetId, startCell}` 或 `{newSheetName,startCell}` 二选一。mode为values或formulas，默认values；formula kind必须在显式formulas模式，不按string前缀推断。百分比/日期支持在values模式保留声明格式。

不允许默认清除 target 范围之外的数据；输出空矩形需显式报告“无结果”，不当作清空指令。现有区域有内容时 prepare摘要列出会覆盖的范围和数量，交由原生工具确认规则处理。文件不能携带任意 shell命令、跨文档目标或宏执行请求。

结果来源声明只证明关联，不证明程序统计正确；验收要求独立expected结果核对。新图表通过正常apply工具依据回写后的数据范围创建，而不是把任意二进制对象混入交换文件。

## Prepare, immutable staging and apply

1. 在 workspace host 校验文件 URI、scheme/authority、真实路径，拒绝越界、路径穿越、非普通文件和不能映射的资源。按约定只允许所选 workspace 的交换目录，私有文件和其他workspace不被隐式开放。
2. 边读边复制到扩展私有 staging，验证source registry、complete、schema/shape、UTF-8、每格类型、哈希与大小；私有副本冻结供后续使用。不能在每批重新打开用户可改写的结果文件。
3. 核对 sourceRevision 和当前revision相同，生成30分钟有效的planId，列出目标、受影响区域、batchCount、单批大小/历史预算及输出语义。prepare不创建工作表、不写单元格。
4. apply按默认<=1,000个普通单元格组织批次，仍遵守2k展开与64MiB双向历史限制。每批走相同候选/commit。新sheet与首批内容在隔离子阶段验证后一起提交，失败不留空表。导入中间值从私有副本获取，不经过模型上下文。
5. 后续只接受精确cursor、该计划前一commit生成的expectedRevision和唯一transactionId；其他内容修改导致stale。取消/失败保留并报告已完成批次，未提交批次无残留。
6. 所有批次完成后返回整体completed状态和实际目标；AI再读取汇总范围、核查公式或创建图表。程序文件按workspace普通文件保留，私有prepared副本超期/关闭后清理，但被历史/恢复点引用的块不得删除。

## Representative semantic example

用户将 Data!D2 从100改为120而未保存。export 的 sourceRevision=7，D2 raw value为120。Python 分地区求和并生成4行结果，source仍引用该快照。prepare在revision7成功，新Summary表及4行数据在第0批提交至revision8；随后图表操作明确使用revision8，提交至revision9。若Python运行期间文档先变成revision8，则旧结果prepare失败；不能仅改manifest为8或读取最新revision后强行套用。

## Failures

`snapshot_stale`、`source_mismatch`、`schema_unsupported`、`artifact_incomplete`、`artifact_hash_mismatch`、`artifact_limit_exceeded`、`type_mismatch`、`target_invalid`、`path_not_allowed`、`plan_expired`、`cursor_invalid` 加上 [session错误](document-session.md)。任何校验失败不修改文档；多批失败必须含completedBatchCount、最后receipt和下一步说明。
