# 工作区 Python 交换 v1

export 最多 1,000,000 格且数据最多 64 MiB，任一超限均拒绝。
它返回 snapshotId、sourceRevision/stateId、manifest/data 的 URI 与宿主路径、哈希及 shape。
Remote SSH 下使用返回的宿主路径，不从客户端 URI 猜测程序位置。

文件位于所选工作区 `.genoffice/work/<snapshotId>/`，manifest 最后发布。
只有 complete=true 才可作为完整输入。公开 manifest 不授予新来源/版本权限。

## 标准库例子

[analyze.py](../scripts/analyze.py) 可以查看、复制或按任务修改，只使用标准库。
先核对表头：summary 模式要求 Region 与 Sales 列，使用原始数值和 Decimal 汇总。

```bash
python3 /workspace/.github/skills/genoffice-workbook/scripts/analyze.py /workspace/.genoffice/work/实际快照ID/input.manifest.json
```

完整值复制使用 `--mode copy --new-sheet-name FullCopy`。
复制公式需显式 `--allow-formulas`；导入后缓存仍视为未验证，不能据此宣称已算对。
脚本不打开/修改 XLSX、不安装依赖、不调用模型，运行交由原生终端权限处理。

## 记录与类型

```json
{"rowOffset":0,"columnOffset":0,"cells":[{"kind":"string","value":"00123"},{"kind":"number","value":0.25,"numberFormat":"0%"}]}
```

偏移相对于范围左上角，从 0 开始。记录按行/列递增，每格恰好一次，不能缺格、重叠或越界。
单条记录最多 256 KiB；宽行可拆成连续 segment，不代表总列数受此字节数限制。

| kind | 字段/语义 |
| --- | --- |
| blank | value=null，区别于空文本 |
| string | value 原文本，包含空串、00123、以 = 开头的字面文本 |
| boolean | value=true/false |
| number | 有限 value，可附 numberFormat；0.25 + 0% 不变成 25 |
| date | 有限 serial、dateSystem、numberFormat，保留日内小数 |
| error | excelCode，例如 #REF!，不当作 0 或空值 |
| formula | formula、cached 标量或 null、calculationState；缓存与公式独立 |

1900 的 serial 60 保留为 Excel 特殊序数，不伪造 Gregorian 日期。
1904 日期按对应系统解释，导入的 dateSystem 必须匹配目标；不自动加时区或平移序数。
禁止 NaN/Infinity。超过工作簿数值精度的编号用 string。
显示文本不代替原始值，pending/error 缓存不当作确认数值。

## 结果 manifest

```json
{
  "schemaVersion":1,"kind":"workbook-result",
  "source":{"snapshotId":"原快照","sourceManifestSha256":"原manifest摘要","sourceDataSha256":"原数据摘要","sourceRevision":7},
  "output":{"fileName":"result.ndjson","sha256":"结果摘要","bytes":123,"recordCount":5,"encoding":"utf-8","format":"typed-ndjson-v1"},
  "shape":{"rows":5,"columns":2,"cellCount":10},
  "target":{"newSheetName":"Summary","startCell":"A1"},
  "mode":"values","complete":true
}
```

也可用 target `{sheetId,startCell}`；两种目标不能混用或跨文档。
rows*columns=cellCount，空矩形不是清空指令。values 为默认模式；formula kind 需显式 formulas 模式，
string 前缀不推断公式。只写目标矩形，不携带 shell 命令、宏或任意二进制对象。

prepareImport 传 bindingId、expectedRevision、snapshotId、resultManifestUri。
它校验路径、来源、完整性、类型、shape、哈希并创建私有不可变副本，返回 30 分钟有效的
planId、目标、覆盖数量、batchCount、nextCursor、chainExpectedRevision；此时不改文档。
摘要证明来源与完整性，业务统计仍需独立核对。

applyImport 沿原 bindingId/planId，传精确 cursor、chainExpectedRevision、唯一 transactionId。
每批最多 1,000 格、一个撤销步骤；首批新表和数据一起提交，失败不留空表。
同 cursor/同事务重试返回旧 receipt，换事务重放旧 token 或跳 cursor 都拒绝。
导入自身创建新表后继续原计划，不在批间重绑定；外部编辑或 undo 令计划 stale。

完成后重新 context/read，再用新绑定/版本创建图表。
部分完成报告 completedCount/remainingCount。过期/关闭/取消只清私有暂存，
不删除用户脚本/结果，也不删除已提交历史。
