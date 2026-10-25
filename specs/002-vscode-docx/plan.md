# Implementation Plan: 单插件 DOCX 编辑

**Branch**: `feature/vscode-docx` | **Date**: 2026-10-24 | **Spec**: [spec.md](spec.md)

**Input**: `specs/002-vscode-docx/spec.md`。逻辑 feature ID 为 `002-vscode-docx`，不更换已建立的 Git 分支。

## Summary

在现有扩展中注册独立 DOCX provider，保持 `genoffice.genoffice-workbook` 和
`genoffice.workbook` 不变。复用现有 DOCX schema、JSON 转换与块级保存计划，
通过薄适配隔离 Electron。新的宿主会话统一手工、AI、程序回写与历史。
DOCX 有独立 HTML 和延迟加载的宿主模块；XLSX 入口保留。首版不重写 Word
引擎、不提供另一聊天界面、不通过磁盘覆盖绕开未保存状态。

## Technical Context

**Language/Version**: TypeScript 5.9 / Node 22；React 19；Tiptap 3.31。
**Primary Dependencies**: 现有 `docx-engine`、Docs 编辑扩展/转换器、Zod、JSZip；
新增 `prosemirror-collab` 1.3.1，用于待确认手工步骤的确认与重基。
**Storage**: 不变原始 DOCX 字节、版本化 PM JSON、历史快照、收据及私有备份。
**Testing**: Vitest、Playwright、真实 VS Code、ZIP 条目保留及安装包测试。
**Target Platform**: 现有 VS Code 支持平台；交付验证以 Linux 开发和 Windows x64
实际安装为主，未执行平台及真实模型测试单列，不作成功推断。
**Project Type**: 现有 monorepo 的插件增强。
**Performance Goals**: 规格 SC-005：1 MiB/200 段/10 表样例冷打开 ≤10 秒，普通编辑 ≤1 秒。
**Constraints**: 不加载 Electron/AI provider 到编辑器；不接收模型任意 XML；
不重解析保存结果并替换原始基底；不改主题数据；保留 XLSX 标识与行为。
**Scale/Scope**: 首版压缩文件 16 MiB，沿用 ZIP 膨胀限制；当前编辑状态 32 MiB；
历史序列化预算 256 MiB；每次手工提交 ≤1000 步且 ≤1 MiB；每批工具 ≤100 操作；
模型读取 ≤100 块且 ≤256 KiB；程序交换 ≤1000 块且 ≤2 MiB。超过限制明确拒绝。
这些是保护上限，性能验收仍使用 SC-005 的较小样例。

## Constitution Check

宪章仍是未批准模板，注释中的例子不作为政策。执行用户决策和 CLAUDE.md。

| Gate         | Design decision                                                   |
| ------------ | ----------------------------------------------------------------- |
| 同一插件     | 保留 ID、XLSX provider/工具；新增 DOCX provider/工具及独立入口    |
| 宿主权威     | 同一 schema 应用步骤，保护/属性校验后原子提交；原生历史统一       |
| 数据保留     | 复用原始块补丁；不支持对象只读；保存后不更换基底                  |
| 原生 AI      | 只注册原生工具，技能/Python 经安全结构化交换回写                  |
| 浏览器边界   | 只放行经审计的纯子路径，不放开整个 AI/UI 桶依赖                   |
| 主题/本地化  | 复用现有词条；新增词条写齐 21 语言；Chrome 语义 token，文档色不变 |
| 不破坏桌面端 | 三处纯 UI 子路径替换及适配点保留桌面默认；相关回归与实际构建      |
| 验证诚实     | 记录首次失败、复测条件和未执行验收；不以代码完成冒充模型/平台通过 |

Phase 1 复查通过，未引入宪章例外。

## Project Structure

```text
specs/002-vscode-docx/
  spec.md plan.md research.md data-model.md quickstart.md tasks.md
  contracts/document-session.md contracts/native-tools.md
packages/docs-editor/
  src/shared/{protocol,schema,policy}.ts
  src/{conversion,DocsEditor,runtime}.tsx
  tests/
packages/docx-session/
  src/{session,operations,save,backup,exchange}.ts
  tests/
apps/vscode-office/
  src/extension/docx/{provider,document,tools,host}.ts
  src/extension/docx-proxy.ts
  src/webview/{docx.html,docx.tsx,docx-theme.css}
  tests/integration/docx*.test.ts
```

薄适配包通过受测的明确导入复用现有 Docs schema/转换器和样式，不整体搬迁上游
编辑器；三处 UI 桶导入改为纯子路径。所有公开入口维持 host/browser 边界。
纯 DTO、schema 与约束放在 docs-editor/shared；docx-session 只从纯子路径依赖它，
前端不导入 Node 会话包。该方向不形成循环依赖。

## Execution Design

1. 添加纯 UI 子路径并完成真实非空文档的无 DOM schema/保存烟测。
2. 宿主解析一次并建立初始 PM JSON；Map 转显式 entries。浏览器不重复转换。
3. 独立逻辑块 ID 使用宿主分配；来源索引只作不变基底锚点。拆分的新逻辑块清理
   继承源元数据并归一化为新块；已有块不随保存重新编号。
4. 手工使用协作步骤确认；捕获追加事务的步骤。宿主校验并补充规范化步骤，区分
   文档 revision 与逐步递增的协作 version。AI 必须先 flush，再校验读版本。
5. 禁用编辑器自身 UndoRedo；原生历史前 flush 并短暂锁定视图，恢复完整内容根后
   重置协作基线，再恢复编辑；保存、读取、备份也等待组合输入和未确认步骤。
6. 来源锚点、XML、关系、保护、对象等不可信属性在宿主检查。schema 校验并不足够。
   模型只用高层操作和安全投影；含隐藏内容的整块程序/AI 覆盖首版拒绝。
7. 基本列表由宿主管理真实 numbering 定义；缺失标题样式由宿主管理 style upserts。
8. 高级表格、修订、锁定 SDT 等设置只读边界，并替换会自行启用 DOM 内编辑的
   NodeView。高级对象不可通过 Backspace、整块替换或伪造步骤绕过保护。
9. 保存复用 `saveDocx` 与现有文件基线、Windows 路径兼容及发布原语；比较未涉及
   ZIP 条目的内容字节，不承诺编辑后整个 ZIP 字节相同。
10. 程序交换只输出安全块投影；受工作区路径和当前版本约束，预览与提交共用收据。
11. 扩展宿主另行构建 DOCX CJS 模块，通过代理延迟加载；Vite 保留 XLSX index.html，
    增加 docx.html，打包显式纳入新的模块与页面资源。

## Validation Sequence

先验证 engine/schema 的复用，再做原子事务/保护/跨保存历史测试，随后浏览器手工
与 AI 工具共存；最终跑现有 XLSX 核心回归、桌面受影响构建、真实 VS Code 两格式
共存和安装包测试。Windows 使用已有隔离构建目录及校验过的离线缓存，不再次
依赖故障的公共 DNS。新增公共依赖按锁文件校验后补入缓存。
