# Tasks: 单插件 DOCX 编辑

Dependencies: Setup → foundation → US1 → US2 → US3 → final gates; US4 protections
are part of foundation and must gate all writes. Parallel markers identify only
independent files; shared contracts and schema changes remain sequential.

## Setup and foundation

- [x] T001 Add pure UI exports and replace three barrel imports in packages/ui/package.json and apps/docs/src/renderer/{font-list,editor/protected-render,editor/shape-draw}.ts.
- [x] T002 Add packages/docs-editor and packages/docx-session workspace manifests/configs; pin collab 1.3.1 and update package-lock.json.
- [x] T003 Add meaningful nonempty schema/conversion/save preservation tests in packages/docx-session/tests/preservation.test.ts before adapters.
- [x] T004 Implement shared DOCX schema and existing-converter adapter in packages/docs-editor/src/shared/schema.ts and src/conversion.ts; no DOM or Electron at host import.
- [x] T005 Add stable logical IDs, trusted source attributes, protected-node validation and bounded transaction checks in packages/docs-editor/src/shared/policy.ts; root ≤32 MiB, steps ≤1000/1 MiB.
- [x] T006 Implement immutable original-based session, separate revision/collabVersion, atomic steps/typed ops, idempotent receipts and 256 MiB history budget in packages/docx-session/src/session.ts.

## US1 — Open, edit and save

- [x] T007 Implement document load/presentation and original-based save with file baselines and Windows path helper in packages/docx-session/src/{open,save}.ts; file ≤16 MiB.
- [x] T008 Implement DOCX custom document/provider, save/as/revert and native history in apps/vscode-office/src/extension/docx/{document,provider}.ts.
- [x] T009 Add lazy provider proxy and separate host bundle in apps/vscode-office/src/extension/docx-proxy.ts and esbuild.config.mjs; existing plugin/editor IDs stay unchanged.
- [x] T010 Add separate docx.html entry and DOCX bridge in apps/vscode-office/src/webview/{docx.html,docx.tsx}, preserving XLSX index.html.
- [x] T011 Implement collab acknowledgement, appended transaction steps, flush/IME and host history integration in packages/docs-editor/src/runtime.ts.
- [x] T012 Implement DocsEditor Word-style UI, paragraph/font/alignment controls and 21-locale labels in packages/docs-editor/src/DocsEditor.tsx and apps/docs/src/renderer/i18n/app/.
- [x] T013 Implement true numbering and missing heading style allocation in packages/docx-session/src/operations.ts; both overlays belong in history roots.
- [x] T014 Implement basic rectangular table creation and row/column operations through validated schema in packages/docs-editor/src/DocsEditor.tsx and session operation tests.
- [x] T015 Add Word/VS Code appearance setting, document width/fonts and semantic chrome CSS in apps/vscode-office/src/webview/docx-theme.css; theme must not rewrite document data.

## US4 — Original content preservation

- [x] T016 Override advanced object editors with static read-only NodeViews in packages/docs-editor/src/shared/schema.ts; reject deletion/forged XML/anchors at host.
- [x] T017 Add original ZIP content, pictures, headers/footers, hidden content, protection and unsafe attribute tests in packages/docx-session/tests/{preservation,policy}.test.ts.
- [x] T018 Verify same-plugin independent/lazy editors and safe media/CSP boundaries in apps/vscode-office/tests/integration/docx.test.ts and build boundary checks.

## US2 — Native chat and shared history

- [x] T019 Add safe content/selection projections and versioned binding handling in packages/docx-session/src/{read,bindings}.ts; read ≤100 blocks/256 KiB, no raw PM/XML or private values.
- [x] T020 Add context/read/apply/status native DOCX tools and explicit confirmations in apps/vscode-office/src/extension/docx/tools.ts; ≤100 high-level operations per atomic batch.
- [x] T021 Extend tool manifest generation and shipped skill references in apps/vscode-office/src/extension/tools/manifest.ts and skills/genoffice-workbook/ without changing existing workbook tools.
- [x] T022 Test manual/tool alternation, 20 edits across save, stale versions, duplicate transactions, protected/private targets and independent documents in packages/docx-session/tests/session.test.ts.

## US3 — Programming and recovery

- [x] T023 Implement safe export/prepareImport/applyImport and native tool wiring in packages/docx-session/src/exchange.ts and apps/vscode-office/src/extension/docx/tools.ts; ≤1000 blocks/2 MiB, strict workspace confinement.
- [x] T024 Implement bounded private backup/restore in packages/docx-session/src/backup.ts and DOCX provider; restoration creates a fresh session identity.
- [x] T025 Add program result hash/version/path and crash recovery tests in packages/docx-session/tests/{exchange,backup}.test.ts and native integration.

## Final gates

- [x] T026 Update root scripts, VSIX resource checks, README/status and feature verification records in apps/vscode-office/ and package.json.
- [x] T027 Run new unit/type/boundary gates, affected desktop compatibility and existing XLSX regression; fix failures without weakening assertions.
- [x] T028 Run browser DOCX editing/collab/format/table/theme tests and realistic SC-005 fixture; record actual timing.
- [x] T029 Run real VS Code DOCX/XLSX coexistence, tools/history/save/backup tests in apps/vscode-office/tests/integration/.
- [x] T030 Build Windows same-plugin VSIX and verify actual installed package; record exact hash and unverified real-model/platform gates.

Each story is independently checked by quickstart.md scenarios. T001–T018 form
an editing/preservation slice; native AI and programming complete the authorized
first release. An unchecked gate is unfinished work, not a request to silently
reduce scope. Spec-quality checklist state is not implementation status.
